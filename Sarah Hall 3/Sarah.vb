''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Program:  S.A.R.A.H. (Self Actuated Residential Automated Habitat)
'Function:  Voice interaction module for the Watchdog system.  Reports on house status,
'           House Temperature, and Weather.
'           Senses presence and starts listening only when someone is present.
'Interfaces:  Watchdog Database
'                  - Event_Current_State (Read only)
'                  - Temp_Current_State  (Read only)
'                  - Event_History       (Insert only)
'                  - X10_Control         (Update only)
'             Internet
'                  - Weather Underground
'                  - Yahoo Weather
'                  - MSN Weather
'             Arduino
'                  - Ultrasonic Sensor   (Digital Input)
'                  - AC Plug Relay       (Digital Output)
'                  - LM335 Temp Sensor   (Analog Input)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Strict Off
Option Explicit On

Public Class Sarah
    Const DISTANCE = 16
    Const SLEEPWAIT = 15
    Const SHOWCAMERA = 60
    Const SLEEPTIME = 300

    Dim WithEvents Recognizer As New SpeechRecognitionEngine
    Dim Synthesizer As New SpeechSynthesizer
    Friend WithEvents Event_HistoryTableAdapter As New WatchdogDataSet1TableAdapters.Event_HistoryTableAdapter
    Dim iPrecipToday, iPrecipTomorrow As Integer
    Dim iCurentTemp, iTodayLow, iTodayHigh, iTomorrowLow, iTomorrowHigh As Integer
    Dim strCurrentCondition, strForecast1, strForecast2, strForecast3 As String
    Dim strWearToday, strWearTomorrow, strWaterOff As String
    Dim bDataFeedError, bWebFeedError1, bWebFeedError2, bWebFeedError3 As Boolean

    Dim serialPort1 As System.IO.Ports.SerialPort
    Dim _ComPort As String = "COM5"
    Dim _BaudRate As Integer = 9600

    Dim bMute, bMonitoring As Boolean
    Dim iSleepCounter As Integer
    Dim tSleep As System.Timers.Timer
    Dim tCheckPresence As System.Timers.Timer
    Dim iOldDistance As Integer
    Dim strImgToday, strImgTomorrow, strImgCurrent As String
    Dim bDoorBellPressed As Boolean


    Private Sub Sarah_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Recognizer.RecognizeAsyncCancel()
        Event_HistoryTableAdapter.InsertQuery("9006", Now)

    End Sub

    Private Sub Sarah_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'WatchdogDataSet1.Insteon_Control' table. You can move, or remove it, as needed.
        Me.Insteon_ControlTableAdapter.Fill(Me.WatchdogDataSet1.Insteon_Control)
        'TODO: This line of code loads data into the 'WatchdogDataSet1.Temp_Control' table. You can move, or remove it, as needed.
        Me.Temp_ControlTableAdapter.Fill(Me.WatchdogDataSet1.Temp_Control)
        'TODO: This line of code loads data into the 'WatchdogDataSet1.Temp_Control' table. You can move, or remove it, as needed.
        Me.Temp_ControlTableAdapter.Fill(Me.WatchdogDataSet1.Temp_Control)
        Dim port As String
        Dim ports As String() = SerialPort.GetPortNames()
        Dim ptPanel As System.Drawing.Point

        Dim badChars As Char() = New Char() {"c"}

        For Each port In ports
            ' .NET Framework has a bug where COM ports are
            ' sometimes appended with a 'c' characeter!
            If port.IndexOfAny(badChars) <> -1 Then
                cbCOMPort.Items.Add(port.Substring(0, port.Length - 1))
            Else
                cbCOMPort.Items.Add(port)
            End If
            cbCOMPort.Text = port
        Next
        If cbCOMPort.Items.Count = 0 Then
            cbCOMPort.Text = ""
        Else
            cbCOMPort.Text = cbCOMPort.Items(0).ToString
        End If

        ptPanel.X = (Me.Width / 2) - (PanelForecast.Width / 2)
        ptPanel.Y = (Me.Height / 2) - (PanelForecast.Height / 2)

        PanelForecast.Location = ptPanel
        PanelForecast.Hide()
        PanelCurrent.Location = ptPanel
        PanelCurrent.Hide()

        'Mute microphone
        bMute = True
        bDoorBellPressed = False

        'Create Thread-safe timer
        tSleep = New System.Timers.Timer(1000)
        AddHandler tSleep.Elapsed, AddressOf tSleep_Elapsed

        tCheckPresence = New System.Timers.Timer(500)
        AddHandler tCheckPresence.Elapsed, AddressOf tCheckPresence_Elapsed

        'SARAH Initialize
        Event_HistoryTableAdapter.InsertQuery("9005", Now)
        InitializeSpeechRecognitionEngine()
        Recognizer.SetInputToDefaultAudioDevice()
        bDataFeedError = False
        UpdateHouse()
        UpdateTemp()
        UpdateWeather()

        lblStatus.Text = "Muted"

        Sarah_Connect()

        Start_Monitor()

    End Sub

#Region "Arduino Code"
    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        lblConnectionStatus.Text = "Connecting"

        If cbCOMPort.Text = "" Then
            MsgBox("Connect Error")
            lblConnectionStatus.Text = "Not Connected"
            Exit Sub
        End If

        Try
            serialPort1 = Nothing
            Dim components As System.ComponentModel.IContainer = New System.ComponentModel.Container()
            serialPort1 = New System.IO.Ports.SerialPort(components)
            serialPort1.PortName = cbCOMPort.Text
            serialPort1.BaudRate = _BaudRate
            serialPort1.ReceivedBytesThreshold = 1

            serialPort1.Open()

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            btnStart.Enabled = True
            cbCOMPort.Enabled = False

            lblConnectionStatus.Text = "Connected"
        Catch ex As Exception
            MsgBox("Connect Error")
            lblConnectionStatus.Text = "Not Connected"
            Exit Sub
        End Try

    End Sub

    Sub Sarah_Connect()
        lblConnectionStatus.Text = "Connecting"

        Try
            serialPort1 = Nothing
            Dim components As System.ComponentModel.IContainer = New System.ComponentModel.Container()
            serialPort1 = New System.IO.Ports.SerialPort(components)
            serialPort1.PortName = _ComPort
            serialPort1.BaudRate = _BaudRate
            serialPort1.ReceivedBytesThreshold = 1

            serialPort1.Open()

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            btnStart.Enabled = True
            cbCOMPort.Enabled = False

            lblConnectionStatus.Text = "Connected"
        Catch ex As Exception
            MsgBox("Connect Error")
            lblConnectionStatus.Text = "Not Connected"
            Exit Sub
        End Try
    End Sub

    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        Stop_Monitor()
        serialPort1.Close()
        tSleep.Enabled = False

        btnConnect.Enabled = True
        btnDisconnect.Enabled = False
        btnStart.Enabled = False
        btnStop.Enabled = False
        cbCOMPort.Enabled = True

        lblConnectionStatus.Text = "Not Connected"
        bMonitoring = False
        'tCheckPresence.SynchronizingObject = Me
        'tCheckPresence.Stop()

    End Sub

    Private Sub Show_Camera()
        Mute()
        Stop_Monitor()
        'tSleep.SynchronizingObject = Me
        'tSleep.Enabled = False
        bDoorBellPressed = True

        'Turn on the monitor
        serialPort1.Write("1")

        'Hide SARAH for x seconds
        Me.Hide()
        iSleepCounter = SHOWCAMERA
        tHideWindow.Start()
    End Sub

    Private Sub GoToSleep()
        Mute()
        Stop_Monitor()
        Event_HistoryTableAdapter.InsertQuery("9008", Now)
        'tSleep.SynchronizingObject = Me
        'tSleep.Enabled = False

        'Turn off the monitor
        serialPort1.Write("0")

        'Mute SARAH for 5 min
        iSleepCounter = SLEEPTIME
        tHideWindow.Start()
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Start_Monitor()
    End Sub

    Sub Start_Monitor()
        serialPort1.DiscardInBuffer()
        tCheckPresence.SynchronizingObject = Me
        tCheckPresence.Start()
        bMonitoring = True
        lblMonitor.Text = "Monitoring"
        btnStop.Enabled = True
        btnStart.Enabled = False
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        Stop_Monitor()
    End Sub

    Sub Stop_Monitor()
        'tCheckPresence.SynchronizingObject = Me
        'tCheckPresence.Stop()
        bMonitoring = False
        lblMonitor.Text = "Not Monitoring"
        btnStart.Enabled = True
        btnStop.Enabled = False
    End Sub

#End Region

#Region "SARAH Code"

    Private Sub InitializeSpeechRecognitionEngine()
        Dim customGrammar As Grammar
        Dim gBuilder As New GrammarBuilder
        Dim myCommands As New Choices

        myCommands.Add("Sarah")
        myCommands.Add("Thanks")
        myCommands.Add("Thanks Sarah")
        myCommands.Add("Thank you")
        myCommands.Add("Thank you Sarah")
        myCommands.Add("Good night")
        myCommands.Add("Good night Sarah")
        myCommands.Add("Garage")
        myCommands.Add("Garage Door")
        myCommands.Add("Mud Room")
        myCommands.Add("Kitchen")
        myCommands.Add("Back Door")
        myCommands.Add("Great Room")
        myCommands.Add("Play Room")
        myCommands.Add("Family Room")
        myCommands.Add("Front Door")
        myCommands.Add("Dining Room")
        myCommands.Add("Master Bath")
        myCommands.Add("Master Bedroom")
        myCommands.Add("Sydneys room")
        myCommands.Add("Ethans room")
        myCommands.Add("Spare Bedroom")
        myCommands.Add("Mailbox")
        myCommands.Add("Mail")
        'myCommands.Add("Cars")
        'myCommands.Add("Car")
        myCommands.Add("House")
        myCommands.Add("House Status")
        myCommands.Add("Upstairs")
        myCommands.Add("Upstairs Status")
        myCommands.Add("Main Level")
        myCommands.Add("Main Level Status")
        myCommands.Add("Basement")
        myCommands.Add("Basement Status")
        myCommands.Add("House Temperature")
        myCommands.Add("Upstairs Temperature")
        myCommands.Add("Main Level Temperature")
        myCommands.Add("Server Room Temperature")
        myCommands.Add("Garage Temperature")
        myCommands.Add("What can I say")
        myCommands.Add("What are you")
        myCommands.Add("Who are you")
        myCommands.Add("Watchdog Status")
        myCommands.Add("Watchdog")
        myCommands.Add("Current Weather")
        myCommands.Add("Whats the weather like")
        myCommands.Add("Forecast")
        myCommands.Add("Weather Forecast")
        myCommands.Add("Weather Today")
        myCommands.Add("What should I wear")
        myCommands.Add("What should I wear today")
        myCommands.Add("What should I wear tomorrow")
        myCommands.Add("Should I turn the water off")
        myCommands.Add("Should I turn off the water")
        myCommands.Add("What is the time")
        myCommands.Add("What time is it")
        myCommands.Add("What is the date")
        myCommands.Add("What day is it")
        myCommands.Add("Light Off")
        myCommands.Add("Light On")
        myCommands.Add("Lights Off")
        myCommands.Add("Lights On")
        myCommands.Add("Temperature Up")
        myCommands.Add("Temperature Down")
        myCommands.Add("What is the air speed velocity of a laden swallow")
        myCommands.Add("Mirror Mirror on the wall")
        myCommands.Add("Im Expecting Company")
        myCommands.Add("Im Expecting Guests")
        myCommands.Add("Go To Sleep")
        myCommands.Add("Show Camra")
        gBuilder.Append(myCommands)

        customGrammar = New Grammar(gBuilder)
        Recognizer.UnloadAllGrammars()
        Recognizer.LoadGrammar(customGrammar)

    End Sub

    Private Sub Recognizer_SpeechRecognized(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs) Handles Recognizer.SpeechRecognized

        If Not bMute Then
            'tSleep.SynchronizingObject = Me
            'tSleep.Stop()
            'tCheckPresence.SynchronizingObject = Me
            'tCheckPresence.Stop()
            lblConfidence.Text = Str(e.Result.Confidence * 100)
            lblDuration.Text = e.Result.Audio.Duration.TotalMilliseconds.ToString
            lblLastInteraction.Text = Now.ToString

            If e.Result.Confidence * 100 >= 75 Then   'Be at least 75% sure 
                'IdentifySpeaker
                Event_HistoryTableAdapter.InsertQuery("9007", Now)
                SpeechToAction(e.Result.Text.ToString)
                iSleepCounter = iSleepCounter + 1
            Else
                SpeechToAction("Please Repeat")
            End If

            'tCheckPresence.SynchronizingObject = Me
            'tCheckPresence.Start()
            tSleep.Start()
        End If
    End Sub

    Private Sub SpeechToAction(ByVal strCommand As String)
        Dim strMessage As String

        Recognizer.RecognizeAsyncCancel()
        strMessage = ""

        Select Case UCase(strCommand)
            Case "SARAH"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = ProvideGreeting()
            Case "THANKS", "THANKS SARAH", "THANK YOU", "THANK YOU SARAH"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "You're welcome."
            Case "GOOD NIGHT", "GOOD NIGHT SARAH"
                If Hour(Now) >= 20 Then
                    Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                    strMessage = "Good night.  Sweet dreams."
                End If
            Case "GARAGE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = GarageStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "GARAGE DOOR"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = GarageDoorStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MUD ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = MudroomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "KITCHEN"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = KitchenStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "BACK DOOR"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = BackDoorStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "GREAT ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = GreatRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "PLAY ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = PlayRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "FAMILY ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = FamilyRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "FRONT DOOR"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = FrontDoorStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "DINING ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = DiningRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MASTER BATH"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = MasterBathStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MASTER BEDROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = MasterBedroomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "SYDNEYS ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = SydneyRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "ETHANS ROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = EthanRoomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "SPARE BEDROOM"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = SpareBedroomStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MAILBOX", "MAIL"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = MailboxStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
                ''Case "CARS", "CAR"
                ''Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                ''UpdateHouse()
                ''If Not bDataFeedError Then
                ''strMessage = CarStatus()
                ''Else
                ''strMessage = "Sorry, I am unable to process that request at this time."
                ''End If
            Case "HOUSE", "HOUSE STATUS"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = HouseStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "UPSTAIRS", "UPSTAIRS STATUS"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = UpstairsStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MAIN LEVEL", "MAIN LEVEL STATUS"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = MainLevelStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "BASEMENT", "BASEMENT STATUS"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = BasementStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "HOUSE TEMPERATURE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH2
                UpdateTemp()
                If Not bDataFeedError Then
                    strMessage = HouseTemp()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "MAIN LEVEL TEMPERATURE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH2
                UpdateTemp()
                If Not bDataFeedError Then
                    strMessage = MainLevelTemp()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "UPSTAIRS TEMPERATURE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH2
                UpdateTemp()
                If Not bDataFeedError Then
                    strMessage = UpstairsTemp()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "SERVER ROOM TEMPERATURE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH2
                UpdateTemp()
                If Not bDataFeedError Then
                    strMessage = "The server room temperature is " & Temp_Current_StateDataGridView.Rows(0).Cells(8).Value.ToString & " degrees"
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "GARAGE TEMPERATURE"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH2
                UpdateTemp()
                If Not bDataFeedError Then
                    strMessage = "The garage temperature is " & Temp_Current_StateDataGridView.Rows(0).Cells(7).Value.ToString & " degrees"
                Else
                    strMessage = "Sorry, I am unable to process that request at this time."
                End If
            Case "WHAT CAN I SAY"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "You can ask me about the house status, house temperature and weather."
            Case "WHAT ARE YOU", "WHO ARE YOU"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "I am Sarah.  Sarah stands for: Self Actuated Residential Automated Habitat"
            Case "WATCHDOG STATUS", "WATCHDOG"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH4
                UpdateHouse()
                If Not bDataFeedError Then
                    strMessage = WatchdogStatus()
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "CURRENT WEATHER", "WHATS THE WEATHER LIKE"
                Me.BackgroundImage = Nothing
                UpdateWeather()
                If Not bWebFeedError1 Then
                    PanelCurrent.Visible = True
                    PanelCurrent.Update()
                    strMessage = "Currently it is " & strCurrentCondition
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "WEATHER TODAY"
                Me.BackgroundImage = Nothing
                UpdateWeather()
                If Not bWebFeedError3 Then
                    PanelCurrent.Visible = True
                    PanelCurrent.Update()
                    strMessage = strForecast1
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "FORECAST", "WEATHER FORECAST"
                Me.BackgroundImage = Nothing
                UpdateWeather()
                If Not bWebFeedError3 Then
                    PanelForecast.Visible = True
                    PanelForecast.Update()
                    strMessage = strForecast2 & "  " & strForecast3
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "WHAT SHOULD I WEAR TODAY", "WHAT SHOULD I WEAR"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                UpdateWeather()
                If Not bWebFeedError1 And Not bWebFeedError2 Then
                    strMessage = strWearToday
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "WHAT SHOULD I WEAR TOMORROW"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                UpdateWeather()
                If Not bWebFeedError1 And Not bWebFeedError2 Then
                    strMessage = strWearTomorrow
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "SHOULD I TURN THE WATER OFF", "SHOULD I TURN OFF THE WATER"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                UpdateWeather()
                If Not bWebFeedError1 And Not bWebFeedError2 Then
                    strMessage = strWaterOff
                Else
                    strMessage = "Sorry, I am unable to process that request at this time. Please ask again."
                End If
            Case "WHAT IS THE TIME", "WHAT TIME IS IT"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = GetTime(Now)
            Case "WHAT IS THE DATE", "WHAT DAY IS IT"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = GetDate(Now)
            Case "LIGHT OFF", "LIGHTS OFF"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                strMessage = "Turning off Light."
                Light1_Off()
            Case "LIGHT ON", "LIGHTS ON"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                strMessage = "Turning on Light."
                Light1_On()
            Case "TEMPERATURE UP"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                strMessage = "Setting thermostat up one degree."
                RaiseTemp(1)
            Case "TEMPERATURE DOWN"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH3
                strMessage = "Setting thermostat down one degree."
                LowerTemp(1)

            Case "WHAT IS THE AIR SPEED VELOCITY OF A LADEN SWALLOW"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "An African or a European Swallow?"
            Case "MIRROR MIRROR ON THE WALL"
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "You are the fairest of them all!"
            Case "IM EXPECTING COMPANY", "IM EXPECTING GUESTS"
                If Hour(Now) >= 17 Or Hour(Now) <= 9 Then
                    Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                    strMessage = "Ok, thanks.  Turning on outside lights for 30 minutes."
                    OutsideLightsOn()
                Else
                    Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                    strMessage = "Ok, thanks.  I will watch for them to arrive."
                End If
            Case "GO TO SLEEP"
                strMessage = "Entering sleep mode"
                GoToSleep()
            Case "SHOW CAMRA"
                Show_Camera()
            Case "PLEASE REPEAT"
                strMessage = "I dont understand.  Please Repeat."
            Case Else
                Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
                strMessage = "I don't understand " & strCommand
        End Select

        Try
            Synthesizer.Speak(strMessage)
        Catch EX As Exception
            MessageBox.Show(EX.Message)
        End Try

        PanelCurrent.Hide()
        PanelForecast.Hide()
        Me.BackgroundImage = WindowsApplication1.My.Resources.SARAH1
        lblCommand.Text = strCommand
        Recognizer.RecognizeAsync(RecognizeMode.Multiple)
    End Sub

    Private Function ProvideGreeting()
        Dim iRandom, i As Integer

        iRandom = Second(Now)
        i = Int(iRandom / 10)
        iRandom = Int(iRandom - (i * 10))
        ProvideGreeting = ""

        Select Case iRandom
            Case 0, 1, 2, 3, 4
                If Hour(Now) < 12 Then
                    ProvideGreeting = "Good Morning."
                ElseIf Hour(Now) >= 12 And Hour(Now) < 17 Then
                    ProvideGreeting = "Good Afternoon."
                Else
                    ProvideGreeting = "Good Evening."
                End If
            Case 5, 6, 7, 8, 9
                ProvideGreeting = "Hello"

        End Select


    End Function

#Region "Garage Status"

    Private Function GarageStatus()
        Dim iGarageDoor, iGarageWindowS, iGarageWindowN, iGarageODoor, iGarageODoorLock, iTemp As Integer
        Dim iCarE, iCarW As Integer

        iGarageDoor = Event_Current_StateDataGridView.Rows(0).Cells(1).Value
        iGarageWindowS = Event_Current_StateDataGridView.Rows(0).Cells(2).Value
        iGarageWindowN = Event_Current_StateDataGridView.Rows(0).Cells(3).Value
        iGarageODoor = Event_Current_StateDataGridView.Rows(0).Cells(4).Value
        iGarageODoorLock = Event_Current_StateDataGridView.Rows(0).Cells(5).Value
        iCarE = Event_Current_StateDataGridView.Rows(0).Cells(37).Value
        iCarW = Event_Current_StateDataGridView.Rows(0).Cells(38).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(7).Value

        If iGarageDoor = 1 And iGarageWindowS = 1 And iGarageWindowN = 1 And iGarageODoor = 1 And iGarageODoorLock = 1 Then
            GarageStatus = "All garage doors, windows and exterior locks are closed"
        Else
            GarageStatus = "Garage Status.  "
            If iGarageDoor = 0 Then
                GarageStatus = GarageStatus & "The garage door is open.  "
            End If
            If iGarageWindowS = 0 Then
                GarageStatus = GarageStatus & "The south window is open.  "
            End If
            If iGarageWindowN = 0 Then
                GarageStatus = GarageStatus & "The north window is open.  "
            End If
            If iGarageODoor = 0 Then
                GarageStatus = GarageStatus & "The outside door is open.  "
            End If
            If iGarageODoorLock = 0 Then
                GarageStatus = GarageStatus & "The outside door is unlocked.  "
            End If
        End If


        If iCarE = 1 And iCarW = 1 Then
            GarageStatus = GarageStatus & "Both cars are in the garage.  "
        End If

        If iCarE = 0 And iCarW = 0 Then
            GarageStatus = GarageStatus & "Neither car is in the garage.  "
        End If

        If iCarE = 0 And iCarW = 1 Then
            GarageStatus = GarageStatus & "Dads car is gone.  "
        End If

        If iCarE = 1 And iCarW = 0 Then
            GarageStatus = GarageStatus & "Moms car is gone.  "
        End If

        GarageStatus = GarageStatus & "And the temperature is " & Str(iTemp) & " degrees"

    End Function

    Private Function GarageDoorStatus()
        Dim iDoor As Integer

        iDoor = Event_Current_StateDataGridView.Rows(0).Cells(1).Value

        GarageDoorStatus = ""
        If iDoor = 0 Then
            GarageDoorStatus = "The Garage Door is Open."
        Else
            GarageDoorStatus = "The Garage Door is Closed."
        End If

    End Function


    Private Function CarStatus()
        Dim iCarE, iCarW As Integer

        iCarE = Event_Current_StateDataGridView.Rows(0).Cells(37).Value
        iCarW = Event_Current_StateDataGridView.Rows(0).Cells(38).Value

        CarStatus = ""

        If iCarE = 1 And iCarW = 1 Then
            CarStatus = "Both cars are in the garage"
        End If

        If iCarE = 0 And iCarW = 0 Then
            CarStatus = "Neither car is in the garage"
        End If

        If iCarE = 0 And iCarW = 1 Then
            CarStatus = "Dads car is gone"
        End If

        If iCarE = 1 And iCarW = 0 Then
            CarStatus = "Moms car is gone"
        End If

    End Function
#End Region

#Region "Upstairs Status"

    Private Function UpstairsStatus()
        Dim iMBWindow, iMBRWindowS, iMBRWindowES, iMBRWindowEN, iMBRTemp, iSWindow, iSTemp, iEWindow, iETemp, iSBRWindowN, iSBRWindowS As Integer
        Dim iUpstairs As Integer

        iMBWindow = Event_Current_StateDataGridView.Rows(0).Cells(23).Value
        iMBRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(24).Value
        iMBRWindowES = Event_Current_StateDataGridView.Rows(0).Cells(25).Value
        iMBRWindowEN = Event_Current_StateDataGridView.Rows(0).Cells(26).Value
        iMBRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(1).Value
        iSWindow = Event_Current_StateDataGridView.Rows(0).Cells(27).Value
        iSTemp = Temp_Current_StateDataGridView.Rows(0).Cells(3).Value
        iEWindow = Event_Current_StateDataGridView.Rows(0).Cells(28).Value
        iETemp = Temp_Current_StateDataGridView.Rows(0).Cells(2).Value
        iSBRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(29).Value
        iSBRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(30).Value

        UpstairsStatus = ""

        If iMBWindow = 1 And iMBRWindowS = 1 And iMBRWindowES = 1 And iMBRWindowEN = 1 And iSWindow = 1 And iEWindow = 1 And iSBRWindowN = 1 And iSBRWindowS = 1 Then
            UpstairsStatus = "All upstairs windows are closed.  "
        Else
            UpstairsStatus = "The upstairs windows are not secure.  "
            If iMBWindow = 0 Then
                UpstairsStatus = UpstairsStatus & "The master bath window is open.  "
            End If
            If iMBRWindowS = 0 Then
                UpstairsStatus = UpstairsStatus & "The master bedroom south side window is open.  "
            End If
            If iMBRWindowES = 0 Then
                UpstairsStatus = UpstairsStatus & "The master bedroom back south window is open.  "
            End If
            If iMBRWindowEN = 0 Then
                UpstairsStatus = UpstairsStatus & "The master bedroom back north window is open.  "
            End If
            If iSWindow = 0 Then
                UpstairsStatus = UpstairsStatus & "Sydneys window is open.  "
            End If
            If iEWindow = 0 Then
                UpstairsStatus = UpstairsStatus & "Ethans window is open.  "
            End If
            If iSBRWindowN = 0 Then
                UpstairsStatus = UpstairsStatus & "The spare bedroom north window is open.  "
            End If
            If iSBRWindowS = 0 Then
                UpstairsStatus = UpstairsStatus & "The spare bedroom south window is open.  "
            End If
        End If

        iUpstairs = iMBRTemp + iETemp + iSTemp
        iUpstairs = iUpstairs / 3

        UpstairsStatus = UpstairsStatus & "and the average upstairs temperature is " & Str(iUpstairs) & " degrees"

    End Function


    Private Function MasterBathStatus()
        Dim iWindow As Integer

        iWindow = Event_Current_StateDataGridView.Rows(0).Cells(22).Value

        MasterBathStatus = ""
        If iWindow = 0 Then
            MasterBathStatus = "The Master bath Window is Open."
        Else
            MasterBathStatus = "The Master bath Window is Closed."
        End If
    End Function

    Private Function MasterBedroomStatus()
        Dim iWindowS, iWindowES, iWindowEN, iTemp As Integer

        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(23).Value
        iWindowES = Event_Current_StateDataGridView.Rows(0).Cells(24).Value
        iWindowEN = Event_Current_StateDataGridView.Rows(0).Cells(25).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(0).Value

        MasterBedroomStatus = ""

        If iWindowS = 1 And iWindowES = 1 And iWindowEN = 1 Then
            MasterBedroomStatus = "All Master Bedroom windows are closed.  "
        Else
            MasterBedroomStatus = "The Master bedroom is not secure.  "
            If iWindowS = 0 Then
                MasterBedroomStatus = MasterBedroomStatus & "The south side window is open.  "
            End If
            If iWindowES = 0 Then
                MasterBedroomStatus = MasterBedroomStatus & "The back side south window is open.  "
            End If
            If iWindowEN = 0 Then
                MasterBedroomStatus = MasterBedroomStatus & "The back side north window is open.  "
            End If
        End If

        MasterBedroomStatus = MasterBedroomStatus & "and the temperature is " & Str(iTemp) & " degrees."

    End Function

    Private Function SydneyRoomStatus()
        Dim iWindow, iTemp As Integer

        iWindow = Event_Current_StateDataGridView.Rows(0).Cells(26).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(2).Value

        If iWindow = 0 Then
            SydneyRoomStatus = "Sydneys Window is Open and the temperature is " & Str(iTemp) & " degrees"
        Else
            SydneyRoomStatus = "Sydneys Window is Closed and the temperature is " & Str(iTemp) & " degrees"
        End If
    End Function

    Private Function EthanRoomStatus()
        Dim iWindow, iTemp As Integer

        iWindow = Event_Current_StateDataGridView.Rows(0).Cells(27).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(1).Value

        If iWindow = 0 Then
            EthanRoomStatus = "Ethans Window is Open and the temperature is " & Str(iTemp) & " degrees"
        Else
            EthanRoomStatus = "Ethans Window is Closed and the temperature is " & Str(iTemp) & " degrees"
        End If
    End Function

    Private Function SpareBedroomStatus()
        Dim iWindowN, iWindowS As Integer

        iWindowN = Event_Current_StateDataGridView.Rows(0).Cells(28).Value
        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(29).Value

        SpareBedroomStatus = ""
        If iWindowN = 1 And iWindowS = 1 Then
            SpareBedroomStatus = "Both Spare Bedroom windows are closed."
        ElseIf iWindowN = 0 And iWindowS = 0 Then
            SpareBedroomStatus = "Both spare bedroom windows are open."
        Else
            If iWindowN = 0 Then
                SpareBedroomStatus = "The Spare Bedroom north window is open."
            End If
            If iWindowS = 0 Then
                SpareBedroomStatus = "The Spare Bedroom south window is open."
            End If
        End If

    End Function


#End Region

#Region "Main Level Status"

    Private Function MainLevelStatus()
        Dim iMudroomDoor, iMudroomDoorLock, iKWindow, iBackDoor, iBackDoorLock, iGRWindowN, iGRWindowS, iPRWindowN, iPRWindowS, iFRWindowS, iFRWindowWS, iFRWindowWC, iFRWindowWN, iFrontDoor, iFrontDoorLock, iDRWindowS, iDRWindowN As Integer
        Dim iKTemp, iPRTemp, iLRTemp, iMain As Integer
        Dim iGarageDoor, iGarageWindowS, iGarageWindowN, iGarageODoor, iGarageODoorLock As Integer

        iMudroomDoor = Event_Current_StateDataGridView.Rows(0).Cells(6).Value
        iMudroomDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(7).Value
        iKWindow = Event_Current_StateDataGridView.Rows(0).Cells(8).Value
        iBackDoor = Event_Current_StateDataGridView.Rows(0).Cells(9).Value
        iBackDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(10).Value
        iGRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(11).Value
        iGRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(12).Value
        iPRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(13).Value
        iPRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(14).Value
        iFRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(15).Value
        iFRWindowWS = Event_Current_StateDataGridView.Rows(0).Cells(16).Value
        iFRWindowWC = Event_Current_StateDataGridView.Rows(0).Cells(17).Value
        iFRWindowWN = Event_Current_StateDataGridView.Rows(0).Cells(18).Value
        iFrontDoor = Event_Current_StateDataGridView.Rows(0).Cells(19).Value
        iFrontDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(20).Value
        iDRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(21).Value
        iDRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(22).Value

        iGarageDoor = Event_Current_StateDataGridView.Rows(0).Cells(1).Value
        iGarageWindowS = Event_Current_StateDataGridView.Rows(0).Cells(2).Value
        iGarageWindowN = Event_Current_StateDataGridView.Rows(0).Cells(3).Value
        iGarageODoor = Event_Current_StateDataGridView.Rows(0).Cells(4).Value
        iGarageODoorLock = Event_Current_StateDataGridView.Rows(0).Cells(5).Value

        iPRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(4).Value
        iKTemp = Temp_Current_StateDataGridView.Rows(0).Cells(5).Value
        iLRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(6).Value

        MainLevelStatus = ""

        If iMudroomDoor = 1 And iKWindow = 1 And iBackDoor = 1 And iBackDoorLock = 1 And iGRWindowN = 1 And iGRWindowS = 1 And iPRWindowN = 1 And iPRWindowS = 1 And iFRWindowS = 1 And iFRWindowWS = 1 And iFRWindowWC = 1 And iFRWindowWN = 1 And iFrontDoor = 1 And iFrontDoorLock = 1 And iDRWindowS = 1 And iDRWindowN = 1 And _
           iGarageDoor = 1 And iGarageODoor = 1 And iGarageODoorLock = 1 And iGarageWindowN = 1 And iGarageWindowS = 1 Then
            MainLevelStatus = "All main level windows, doors and exterior locks are closed.  "
        Else
            MainLevelStatus = "Main level status.  "
            If iMudroomDoor = 0 Then
                MainLevelStatus = MainLevelStatus & "The mud room door is open.  "
            End If
            If iMudroomDoorLock = 0 Then
                MainLevelStatus = MainLevelStatus & "The mud room door is unlocked.  "
            End If
            If iKWindow = 0 Then
                MainLevelStatus = MainLevelStatus & "The kitchen window is open.  "
            End If
            If iBackDoor = 0 Then
                MainLevelStatus = MainLevelStatus & "The back door is open.  "
            End If
            If iBackDoorLock = 0 Then
                MainLevelStatus = MainLevelStatus & "The back door is unlocked.  "
            End If
            If iGRWindowN = 0 Then
                MainLevelStatus = MainLevelStatus & "The great room north window is open.  "
            End If
            If iGRWindowS = 0 Then
                MainLevelStatus = MainLevelStatus & "The great room south window is open.  "
            End If
            If iPRWindowN = 0 Then
                MainLevelStatus = MainLevelStatus & "The play room north window is open.  "
            End If
            If iPRWindowS = 0 Then
                MainLevelStatus = MainLevelStatus & "The play room south window is open.  "
            End If
            If iFRWindowS = 0 Then
                MainLevelStatus = MainLevelStatus & "The family room south window is open.  "
            End If
            If iFRWindowWS = 0 Then
                MainLevelStatus = MainLevelStatus & "The family room south bay window is open.  "
            End If
            If iFRWindowWC = 0 Then
                MainLevelStatus = MainLevelStatus & "The family room center bay window is open.  "
            End If
            If iFRWindowWN = 0 Then
                MainLevelStatus = MainLevelStatus & "The family room north bay window is open.  "
            End If
            If iFrontDoor = 0 Then
                MainLevelStatus = MainLevelStatus & "The front door is open.  "
            End If
            If iFrontDoorLock = 0 Then
                MainLevelStatus = MainLevelStatus & "The front door is unlocked.  "
            End If
            If iDRWindowS = 0 Then
                MainLevelStatus = MainLevelStatus & "The dining room south window is open.  "
            End If
            If iDRWindowN = 0 Then
                MainLevelStatus = MainLevelStatus & "The dining room north window is open.  "
            End If
            If iGarageDoor = 0 Then
                MainLevelStatus = MainLevelStatus & "The garage door is open.  "
            End If
            If iGarageWindowS = 0 Then
                MainLevelStatus = MainLevelStatus & "The garage south window is open.  "
            End If
            If iGarageWindowN = 0 Then
                MainLevelStatus = MainLevelStatus & "The garage north window is open.  "
            End If
            If iGarageODoor = 0 Then
                MainLevelStatus = MainLevelStatus & "The garage outside door is open.  "
            End If
            If iGarageODoorLock = 0 Then
                MainLevelStatus = MainLevelStatus & "The garage outside door is unlocked.  "
            End If

        End If


        iMain = iPRTemp + iKTemp + iLRTemp
        iMain = iMain / 3

        MainLevelStatus = MainLevelStatus & "and the average main level temperature is " & Str(iMain) & " degrees"

    End Function

    Private Function MudroomStatus()
        Dim iMudroomDoor, iMudroomDoorLock As Integer

        iMudroomDoor = Event_Current_StateDataGridView.Rows(0).Cells(6).Value
        iMudroomDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(7).Value

        MudroomStatus = ""

        If iMudroomDoor = 1 And iMudroomDoorLock = 1 Then
            MudroomStatus = "The Mud room door is closed and locked"
        End If

        If iMudroomDoor = 1 And iMudroomDoorLock = 0 Then
            MudroomStatus = "The Mud room door is closed but not locked"
        End If

        If iMudroomDoor = 0 Then
            MudroomStatus = "The Mud room door is open"
        End If

    End Function

    Private Function KitchenStatus()
        Dim iWindow, iDoor, iDoorLock, iTemp As Integer

        iWindow = Event_Current_StateDataGridView.Rows(0).Cells(8).Value
        iDoor = Event_Current_StateDataGridView.Rows(0).Cells(9).Value
        iDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(10).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(5).Value

        If iWindow = 0 Then
            KitchenStatus = "The Kitchen Window is Open.  "
        Else
            KitchenStatus = "The Kitchen Window is Closed.  "
        End If

        If iDoor = 1 And iDoorLock = 1 Then
            KitchenStatus = KitchenStatus & "The Back door is closed and locked.  "
        End If

        If iDoor = 1 And iDoorLock = 0 Then
            KitchenStatus = KitchenStatus & "The Back door is closed but not locked.  "
        End If

        If iDoor = 0 Then
            KitchenStatus = KitchenStatus & "The Back door is open.  "
        End If

        KitchenStatus = KitchenStatus & "And the temperature is " & Str(iTemp) & " degrees"


    End Function

    Private Function BackDoorStatus()
        Dim iDoor, iDoorLock As Integer

        iDoor = Event_Current_StateDataGridView.Rows(0).Cells(9).Value
        iDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(10).Value

        BackDoorStatus = ""

        If iDoor = 1 And iDoorLock = 1 Then
            BackDoorStatus = "The Back door is closed and locked"
        End If

        If iDoor = 1 And iDoorLock = 0 Then
            BackDoorStatus = "The Back door is closed but not locked"
        End If

        If iDoor = 0 Then
            BackDoorStatus = "The Back door is open"
        End If

    End Function

    Private Function GreatRoomStatus()
        Dim iWindowN, iWindowS As Integer

        iWindowN = Event_Current_StateDataGridView.Rows(0).Cells(11).Value
        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(12).Value

        GreatRoomStatus = ""

        If iWindowN = 1 And iWindowS = 1 Then
            GreatRoomStatus = "Both great room windows are closed"
        End If

        If iWindowN = 0 And iWindowS = 0 Then
            GreatRoomStatus = "Both great room windows are open"
        End If

        If iWindowN = 0 And iWindowS = 1 Then
            GreatRoomStatus = "The great room north window is open"
        End If

        If iWindowN = 1 And iWindowS = 0 Then
            GreatRoomStatus = "The great room south window is open"
        End If

    End Function

    Private Function PlayRoomStatus()
        Dim iWindowN, iWindowS As Integer

        iWindowN = Event_Current_StateDataGridView.Rows(0).Cells(13).Value
        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(14).Value

        PlayRoomStatus = ""

        If iWindowN = 1 And iWindowS = 1 Then
            PlayRoomStatus = "Both play room windows are closed"
        End If

        If iWindowN = 0 And iWindowS = 0 Then
            PlayRoomStatus = "Both play room windows are open"
        End If

        If iWindowN = 0 And iWindowS = 1 Then
            PlayRoomStatus = "The play room north window is open"
        End If

        If iWindowN = 1 And iWindowS = 0 Then
            PlayRoomStatus = "The play room south window is open"
        End If

    End Function

    Private Function FamilyRoomStatus()
        Dim iWindowS, iWindowWS, iWindowWC, iWindowWN, iTemp As Integer

        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(15).Value
        iWindowWS = Event_Current_StateDataGridView.Rows(0).Cells(16).Value
        iWindowWC = Event_Current_StateDataGridView.Rows(0).Cells(17).Value
        iWindowWN = Event_Current_StateDataGridView.Rows(0).Cells(18).Value
        iTemp = Temp_Current_StateDataGridView.Rows(0).Cells(6).Value

        FamilyRoomStatus = ""

        If iWindowS = 1 And iWindowWS = 1 And iWindowWC = 1 And iWindowWN = 1 Then
            FamilyRoomStatus = "All Family Room windows are closed.  "
        Else
            FamilyRoomStatus = "Family room status.  "
            If iWindowS = 0 Then
                FamilyRoomStatus = FamilyRoomStatus & "The south window is open.  "
            End If
            If iWindowWS = 0 Then
                FamilyRoomStatus = FamilyRoomStatus & "The South bay window is open.  "
            End If
            If iWindowWC = 0 Then
                FamilyRoomStatus = FamilyRoomStatus & "The Center bay window is open.  "
            End If
            If iWindowWN = 0 Then
                FamilyRoomStatus = FamilyRoomStatus & "The North bay window is open.  "
            End If
        End If

        FamilyRoomStatus = FamilyRoomStatus & "and the temperature is " & Str(iTemp) & " degrees."

    End Function

    Private Function FrontDoorStatus()
        Dim iFrontDoor, iFrontDoorLock As Integer

        iFrontDoor = Event_Current_StateDataGridView.Rows(0).Cells(19).Value
        iFrontDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(20).Value

        FrontDoorStatus = ""

        If iFrontDoor = 1 And iFrontDoorLock = 1 Then
            FrontDoorStatus = "The front door is closed and locked"
        End If

        If iFrontDoor = 1 And iFrontDoorLock = 0 Then
            FrontDoorStatus = "The front door is closed but not locked"
        End If

        If iFrontDoor = 0 Then
            FrontDoorStatus = "The front door is open"
        End If

    End Function

    Private Function DiningRoomStatus()
        Dim iWindowN, iWindowS As Integer

        iWindowS = Event_Current_StateDataGridView.Rows(0).Cells(21).Value
        iWindowN = Event_Current_StateDataGridView.Rows(0).Cells(22).Value

        DiningRoomStatus = ""

        If iWindowN = 1 And iWindowS = 1 Then
            DiningRoomStatus = "Both dining room windows are closed"
        End If

        If iWindowN = 0 And iWindowS = 0 Then
            DiningRoomStatus = "Both dining room windows are open"
        End If

        If iWindowN = 0 And iWindowS = 1 Then
            DiningRoomStatus = "The dining room north window is open"
        End If

        If iWindowN = 1 And iWindowS = 0 Then
            DiningRoomStatus = "The dining room south window is open"
        End If

    End Function

#End Region

#Region "Basement Status"

    Private Function BasementStatus()
        Dim iBDoor, iBDoorLock, iBWindowS, iBWindowN, iBStorageDoor, iSRTemp As Integer

        iBDoor = Event_Current_StateDataGridView.Rows(0).Cells(32).Value
        iBDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(33).Value
        iBWindowS = Event_Current_StateDataGridView.Rows(0).Cells(34).Value
        iBWindowN = Event_Current_StateDataGridView.Rows(0).Cells(35).Value
        iBStorageDoor = Event_Current_StateDataGridView.Rows(0).Cells(36).Value
        iSRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(8).Value

        BasementStatus = ""

        If iBDoor = 1 And iBDoorLock = 1 And iBWindowS = 1 And iBWindowN = 1 And iBStorageDoor = 1 Then
            BasementStatus = "All basement windows, doors and locks are closed.  "
        Else
            BasementStatus = "The basement is not secure.  "
            If iBDoor = 0 Then
                BasementStatus = BasementStatus & "The Basement outside door is open.  "
            End If
            If iBDoorLock = 0 Then
                BasementStatus = BasementStatus & "The basement outside door is unlocked.  "
            End If
            If iBWindowS = 0 Then
                BasementStatus = BasementStatus & "The basement south window is open.  "
            End If
            If iBWindowN = 0 Then
                BasementStatus = BasementStatus & "The basement north window is open.  "
            End If
            If iBStorageDoor = 0 Then
                BasementStatus = BasementStatus & "The basement storage door is open.  "
            End If
        End If

        BasementStatus = BasementStatus & "and the server room temperature is " & Str(iSRTemp) & " degrees"

    End Function

#End Region

#Region "Misc Status"

    Private Function HouseStatus()
        Dim iMBWindow, iMBRWindowS, iMBRWindowES, iMBRWindowEN, iSWindow, iEWindow, iSBRWindowN, iSBRWindowS As Integer
        Dim iMudroomDoor, iMudroomDoorLock, iKWindow, iBackDoor, iBackDoorLock, iGRWindowN, iGRWindowS, iPRWindowN, iPRWindowS, iFRWindowS, iFRWindowWS, iFRWindowWC, iFRWindowWN, iFrontDoor, iFrontDoorLock, iDRWindowS, iDRWindowN As Integer
        Dim iBDoor, iBDoorLock, iBWindowS, iBWindowN, iBStorageDoor As Integer
        Dim iHouseTemp, iMBRTemp, iSTemp, iETemp, iKTemp, iPRTemp, iLRTemp As Integer
        Dim iGarageDoor, iGarageWindowS, iGarageWindowN, iGarageODoor, iGarageODoorLock As Integer
        Dim iMailbox As Integer

        iMBWindow = Event_Current_StateDataGridView.Rows(0).Cells(23).Value
        iMBRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(24).Value
        iMBRWindowES = Event_Current_StateDataGridView.Rows(0).Cells(25).Value
        iMBRWindowEN = Event_Current_StateDataGridView.Rows(0).Cells(26).Value
        iMBRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(1).Value
        iSWindow = Event_Current_StateDataGridView.Rows(0).Cells(27).Value
        iSTemp = Temp_Current_StateDataGridView.Rows(0).Cells(3).Value
        iEWindow = Event_Current_StateDataGridView.Rows(0).Cells(28).Value
        iETemp = Temp_Current_StateDataGridView.Rows(0).Cells(2).Value
        iSBRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(29).Value
        iSBRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(30).Value

        iMudroomDoor = Event_Current_StateDataGridView.Rows(0).Cells(6).Value
        iMudroomDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(7).Value
        iKWindow = Event_Current_StateDataGridView.Rows(0).Cells(8).Value
        iBackDoor = Event_Current_StateDataGridView.Rows(0).Cells(9).Value
        iBackDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(10).Value
        iGRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(11).Value
        iGRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(12).Value
        iPRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(13).Value
        iPRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(14).Value
        iFRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(15).Value
        iFRWindowWS = Event_Current_StateDataGridView.Rows(0).Cells(16).Value
        iFRWindowWC = Event_Current_StateDataGridView.Rows(0).Cells(17).Value
        iFRWindowWN = Event_Current_StateDataGridView.Rows(0).Cells(18).Value
        iFrontDoor = Event_Current_StateDataGridView.Rows(0).Cells(19).Value
        iFrontDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(20).Value
        iDRWindowS = Event_Current_StateDataGridView.Rows(0).Cells(21).Value
        iDRWindowN = Event_Current_StateDataGridView.Rows(0).Cells(22).Value
        iPRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(4).Value
        iKTemp = Temp_Current_StateDataGridView.Rows(0).Cells(5).Value
        iLRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(6).Value

        iBDoor = Event_Current_StateDataGridView.Rows(0).Cells(32).Value
        iBDoorLock = Event_Current_StateDataGridView.Rows(0).Cells(33).Value
        iBWindowS = Event_Current_StateDataGridView.Rows(0).Cells(34).Value
        iBWindowN = Event_Current_StateDataGridView.Rows(0).Cells(35).Value
        iBStorageDoor = Event_Current_StateDataGridView.Rows(0).Cells(36).Value

        iGarageDoor = Event_Current_StateDataGridView.Rows(0).Cells(1).Value
        iGarageWindowS = Event_Current_StateDataGridView.Rows(0).Cells(2).Value
        iGarageWindowN = Event_Current_StateDataGridView.Rows(0).Cells(3).Value
        iGarageODoor = Event_Current_StateDataGridView.Rows(0).Cells(4).Value
        iGarageODoorLock = Event_Current_StateDataGridView.Rows(0).Cells(5).Value

        iMailbox = Event_Current_StateDataGridView.Rows(0).Cells(31).Value



        HouseStatus = ""

        If iMBWindow = 1 And iMBRWindowS = 1 And iMBRWindowES = 1 And iMBRWindowEN = 1 And iSWindow = 1 And iEWindow = 1 And iSBRWindowN = 1 And iSBRWindowS = 1 And _
           iMudroomDoor = 1 And iKWindow = 1 And iBackDoor = 1 And iBackDoorLock = 1 And iGRWindowN = 1 And iGRWindowS = 1 And iPRWindowN = 1 And iPRWindowS = 1 And iFRWindowS = 1 And iFRWindowWS = 1 And iFRWindowWC = 1 And iFRWindowWN = 1 And iFrontDoor = 1 And iFrontDoorLock = 1 And iDRWindowS = 1 And iDRWindowN = 1 And _
           iBDoor = 1 And iBDoorLock = 1 And iBWindowS = 1 And iBWindowN = 1 And iBStorageDoor = 1 And iGarageDoor = 1 And iGarageODoor = 1 And iGarageODoorLock = 1 And iGarageWindowN = 1 And iGarageWindowS = 1 Then
            HouseStatus = "All windows, doors and exterior locks are closed.  "
        Else
            HouseStatus = "House Status.  "
            If iMBWindow = 0 Then
                HouseStatus = HouseStatus & "The master bath window is open.  "
            End If
            If iMBRWindowS = 0 Then
                HouseStatus = HouseStatus & "The master bedroom south side window is open.  "
            End If
            If iMBRWindowES = 0 Then
                HouseStatus = HouseStatus & "The master bedroom back south window is open.  "
            End If
            If iMBRWindowEN = 0 Then
                HouseStatus = HouseStatus & "The master bedroom back north window is open.  "
            End If
            If iSWindow = 0 Then
                HouseStatus = HouseStatus & "Sydneys window is open.  "
            End If
            If iEWindow = 0 Then
                HouseStatus = HouseStatus & "Ethans window is open.  "
            End If
            If iSBRWindowN = 0 Then
                HouseStatus = HouseStatus & "The spare bedroom north window is open.  "
            End If
            If iSBRWindowS = 0 Then
                HouseStatus = HouseStatus & "The spare bedroom south window is open.  "
            End If
            If iMudroomDoor = 0 Then
                HouseStatus = HouseStatus & "The mud room door is open.  "
            End If
            If iMudroomDoorLock = 0 Then
                HouseStatus = HouseStatus & "The mud room door is unlocked.  "
            End If
            If iKWindow = 0 Then
                HouseStatus = HouseStatus & "The kitchen window is open.  "
            End If
            If iBackDoor = 0 Then
                HouseStatus = HouseStatus & "The back door is open.  "
            End If
            If iBackDoorLock = 0 Then
                HouseStatus = HouseStatus & "The back door is unlocked.  "
            End If
            If iGRWindowN = 0 Then
                HouseStatus = HouseStatus & "The great room north window is open.  "
            End If
            If iGRWindowS = 0 Then
                HouseStatus = HouseStatus & "The great room south window is open.  "
            End If
            If iPRWindowN = 0 Then
                HouseStatus = HouseStatus & "The play room north window is open.  "
            End If
            If iPRWindowS = 0 Then
                HouseStatus = HouseStatus & "The play room south window is open.  "
            End If
            If iFRWindowS = 0 Then
                HouseStatus = HouseStatus & "The family room south window is open.  "
            End If
            If iFRWindowWS = 0 Then
                HouseStatus = HouseStatus & "The family room south bay window is open.  "
            End If
            If iFRWindowWC = 0 Then
                HouseStatus = HouseStatus & "The family room center bay window is open.  "
            End If
            If iFRWindowWN = 0 Then
                HouseStatus = HouseStatus & "The family room north bay window is open.  "
            End If
            If iFrontDoor = 0 Then
                HouseStatus = HouseStatus & "The front door is open.  "
            End If
            If iFrontDoorLock = 0 Then
                HouseStatus = HouseStatus & "The front door is unlocked.  "
            End If
            If iDRWindowS = 0 Then
                HouseStatus = HouseStatus & "The dining room south window is open.  "
            End If
            If iDRWindowN = 0 Then
                HouseStatus = HouseStatus & "The dining room north window is open.  "
            End If
            If iBDoor = 0 Then
                HouseStatus = HouseStatus & "The Basement outside door is open.  "
            End If
            If iBDoorLock = 0 Then
                HouseStatus = HouseStatus & "The basement outside door is unlocked.  "
            End If
            If iBWindowS = 0 Then
                HouseStatus = HouseStatus & "The basement south window is open.  "
            End If
            If iBWindowN = 0 Then
                HouseStatus = HouseStatus & "The basement north window is open.  "
            End If
            If iBStorageDoor = 0 Then
                HouseStatus = HouseStatus & "The basement storage door is open.  "
            End If
            If iGarageDoor = 0 Then
                HouseStatus = HouseStatus & "The garage door is open.  "
            End If
            If iGarageWindowS = 0 Then
                HouseStatus = HouseStatus & "The garage south window is open.  "
            End If
            If iGarageWindowN = 0 Then
                HouseStatus = HouseStatus & "The garage north window is open.  "
            End If
            If iGarageODoor = 0 Then
                HouseStatus = HouseStatus & "The garage outside door is open.  "
            End If
            If iGarageODoorLock = 0 Then
                HouseStatus = HouseStatus & "The garage outside door is unlocked.  "
            End If

        End If

        iHouseTemp = iPRTemp + iKTemp + iLRTemp + iMBRTemp + iETemp + iSTemp
        iHouseTemp = iHouseTemp / 6

        HouseStatus = HouseStatus & "the average house temperature is " & Str(iHouseTemp) & " degrees  "

        If iMailbox = 0 Then
            HouseStatus = HouseStatus & "and The mail has not been delivered"
        End If

        If iMailbox = 1 Then
            HouseStatus = HouseStatus & "and The mail has been delivered"
        End If

        If iMailbox = 2 Then
            HouseStatus = HouseStatus & "and The mail has been retrieved"
        End If


    End Function

    Private Function WatchdogStatus()
        Dim dtWatchdog, dtNow As Date

        dtWatchdog = Event_Current_StateDataGridView.Rows(0).Cells(39).Value
        dtNow = Now

        WatchdogStatus = ""
        If Year(dtWatchdog) = Year(dtNow) And Month(dtWatchdog) = Month(dtNow) And DatePart(DateInterval.Day, dtWatchdog) = DatePart(DateInterval.Day, dtNow) And Hour(dtWatchdog) = Hour(dtNow) And (Minute(dtWatchdog) >= (Minute(dtNow) - 1) Or (Minute(dtWatchdog) = 0 And Minute(dtNow) = 59)) Then
            WatchdogStatus = "Watchdog is active."
        Else
            WatchdogStatus = "Watchdog is not active.  It stopped running on " & GetDate(dtWatchdog) & " at " & GetTime(dtWatchdog)
        End If

    End Function

    Private Function MailboxStatus()
        Dim iMailbox As Integer

        iMailbox = Event_Current_StateDataGridView.Rows(0).Cells(31).Value

        MailboxStatus = ""

        If iMailbox = 0 Then
            MailboxStatus = "The mail has not been delivered"
        End If

        If iMailbox = 1 Then
            MailboxStatus = "The mail has been delivered"
        End If

        If iMailbox = 2 Then
            MailboxStatus = "The mail has been retrieved"
        End If

    End Function


    Private Function GetTime(ByRef dtNow As Date)
        Dim strAMPM As String
        Dim iHour, iMinute As Integer
        Dim bGetMinute As Boolean

        GetTime = ""
        iHour = Hour(dtNow)

        If iHour >= 12 Then
            strAMPM = "Pee emm"
            iHour = iHour - 12
        Else
            strAMPM = "ae emm("")"
        End If

        Select Case iHour
            Case 1
                GetTime = "One "
            Case 2
                GetTime = "Two "
            Case 3
                GetTime = "Three "
            Case 4
                GetTime = "Four "
            Case 5
                GetTime = "Five "
            Case 6
                GetTime = "Six "
            Case 7
                GetTime = "Seven "
            Case 8
                GetTime = "Eight "
            Case 9
                GetTime = "Nine "
            Case 10
                GetTime = "Ten "
            Case 11
                GetTime = "Eleven "
            Case 0, 12
                GetTime = "Twelve "
        End Select

        bGetMinute = True
        iMinute = Minute(dtNow)
        Select Case iMinute
            Case Is = 0
                'Do nothing
            Case Is < 10
                GetTime = GetTime & "Oh "
            Case Is < 20
                bGetMinute = False
                Select Case iMinute
                    Case Is = 11
                        GetTime = GetTime & "Eleven "
                    Case Is = 12
                        GetTime = GetTime & "Twelve "
                    Case Is = 13
                        GetTime = GetTime & "Thirteen "
                    Case Is = 14
                        GetTime = GetTime & "Fourteen "
                    Case Is = 15
                        GetTime = GetTime & "Fifteen "
                    Case Is = 16
                        GetTime = GetTime & "Sixteen "
                    Case Is = 17
                        GetTime = GetTime & "Seventeen "
                    Case Is = 18
                        GetTime = GetTime & "Eighteen "
                    Case Is = 19
                        GetTime = GetTime & "Nineteen "
                End Select
            Case Is < 30
                GetTime = GetTime & "Twenty "
                iMinute = iMinute - 20
            Case Is < 40
                GetTime = GetTime & "Thirty "
                iMinute = iMinute - 30
            Case Is < 50
                GetTime = GetTime & "Fourty "
                iMinute = iMinute - 40
            Case Else
                GetTime = GetTime & "Fifty "
                iMinute = iMinute - 50
        End Select

        If bGetMinute Then
            Select Case iMinute
                Case Is = 1
                    GetTime = GetTime & "One "
                Case Is = 2
                    GetTime = GetTime & "Two "
                Case Is = 3
                    GetTime = GetTime & "Three "
                Case Is = 4
                    GetTime = GetTime & "Four "
                Case Is = 5
                    GetTime = GetTime & "Five "
                Case Is = 6
                    GetTime = GetTime & "Six "
                Case Is = 7
                    GetTime = GetTime & "Seven "
                Case Is = 8
                    GetTime = GetTime & "Eight "
                Case Is = 9
                    GetTime = GetTime & "Nine "
            End Select
        End If

        GetTime = GetTime & strAMPM

    End Function

    Private Function GetDate(ByVal dtNow As Date)
        Dim iMonth, iDay, iDayofweek As Integer

        GetDate = ""
        iDayofweek = dtNow.DayOfWeek
        Select Case iDayofweek
            Case Is = 0
                GetDate = "Sunday "
            Case Is = 1
                GetDate = "Monday "
            Case Is = 2
                GetDate = "Tuesday "
            Case Is = 3
                GetDate = "Wednesday "
            Case Is = 4
                GetDate = "Thursday "
            Case Is = 5
                GetDate = "Friday "
            Case Is = 6
                GetDate = "Saturday "
        End Select

        iMonth = Month(dtNow)
        Select Case iMonth
            Case Is = 1
                GetDate = GetDate & "January "
            Case Is = 2
                GetDate = GetDate & "February "
            Case Is = 3
                GetDate = GetDate & "March "
            Case Is = 4
                GetDate = GetDate & "April "
            Case Is = 5
                GetDate = GetDate & "May "
            Case Is = 6
                GetDate = GetDate & "June "
            Case Is = 7
                GetDate = GetDate & "July "
            Case Is = 8
                GetDate = GetDate & "August "
            Case Is = 9
                GetDate = GetDate & "September "
            Case Is = 10
                GetDate = GetDate & "October "
            Case Is = 11
                GetDate = GetDate & "November "
            Case Is = 12
                GetDate = GetDate & "December "
        End Select

        iDay = Microsoft.VisualBasic.DateAndTime.Day(dtNow)
        Select Case iDay
            Case Is = 1
                GetDate = GetDate & "First"
            Case Is = 2
                GetDate = GetDate & "Second"
            Case Is = 3
                GetDate = GetDate & "Third"
            Case Is = 4
                GetDate = GetDate & "Fourth"
            Case Is = 5
                GetDate = GetDate & "Fifth"
            Case Is = 6
                GetDate = GetDate & "Sixth"
            Case Is = 7
                GetDate = GetDate & "Seventh"
            Case Is = 8
                GetDate = GetDate & "Eighth"
            Case Is = 9
                GetDate = GetDate & "Nineth"
            Case Is = 10
                GetDate = GetDate & "Tenth"
            Case Is = 11
                GetDate = GetDate & "Eleventh"
            Case Is = 12
                GetDate = GetDate & "Twelvth"
            Case Is = 13
                GetDate = GetDate & "Thirteenth"
            Case Is = 14
                GetDate = GetDate & "Fourteenth"
            Case Is = 15
                GetDate = GetDate & "Fifteenth"
            Case Is = 16
                GetDate = GetDate & "Sixteenth"
            Case Is = 17
                GetDate = GetDate & "Seventeenth"
            Case Is = 18
                GetDate = GetDate & "Eighteenth"
            Case Is = 19
                GetDate = GetDate & "Nineteenth"
            Case Is = 20
                GetDate = GetDate & "Twenth"
            Case Is = 21
                GetDate = GetDate & "Twenty First"
            Case Is = 22
                GetDate = GetDate & "Twenty Second"
            Case Is = 23
                GetDate = GetDate & "Twenty Third"
            Case Is = 24
                GetDate = GetDate & "Twenty Fourth"
            Case Is = 25
                GetDate = GetDate & "Twenty Fifth"
            Case Is = 26
                GetDate = GetDate & "Twenty Sixth"
            Case Is = 27
                GetDate = GetDate & "Twenty Seventh"
            Case Is = 28
                GetDate = GetDate & "Twenty Eighth"
            Case Is = 29
                GetDate = GetDate & "Twenty Nineth"
            Case Is = 30
                GetDate = GetDate & "Thirtyth"
            Case Is = 31
                GetDate = GetDate & "Thirty First"
        End Select

    End Function

#End Region

#Region "Temperature"

    Private Function HouseTemp()
        Dim iMBRTemp, iSRTemp, iERTemp, iPRTemp, iKTemp, iLRTemp, iHouse As Integer

        iPRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(4).Value
        iKTemp = Temp_Current_StateDataGridView.Rows(0).Cells(5).Value
        iLRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(6).Value
        iMBRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(1).Value
        iERTemp = Temp_Current_StateDataGridView.Rows(0).Cells(2).Value
        iSRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(3).Value

        iHouse = iPRTemp + iKTemp + iLRTemp + iMBRTemp + iERTemp + iSRTemp
        iHouse = iHouse / 6

        HouseTemp = "The average house temperature is " & Str(iHouse) & " degrees"
    End Function

    Private Function MainLevelTemp()
        Dim iPRTemp, iKTemp, iLRTemp, iMain As Integer

        iPRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(4).Value
        iKTemp = Temp_Current_StateDataGridView.Rows(0).Cells(5).Value
        iLRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(6).Value

        iMain = iPRTemp + iKTemp + iLRTemp
        iMain = iMain / 3

        MainLevelTemp = "The average main level temperature is " & Str(iMain) & " degrees"
    End Function

    Private Function UpstairsTemp()
        Dim iMBRTemp, iSRTemp, iERTemp, iUpstairs As Integer

        iMBRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(1).Value
        iERTemp = Temp_Current_StateDataGridView.Rows(0).Cells(2).Value
        iSRTemp = Temp_Current_StateDataGridView.Rows(0).Cells(3).Value

        iUpstairs = iMBRTemp + iERTemp + iSRTemp
        iUpstairs = iUpstairs / 3

        UpstairsTemp = "The average upstairs temperature is " & Str(iUpstairs) & " degrees"
    End Function
#End Region

#Region "Weather Update"

    Private Sub UpdateWeather()
        Dim rssFeed1 As HttpWebRequest
        Dim rssFeed2 As HttpWebRequest
        Dim rssData1 As DataSet = New DataSet()
        Dim rssData2 As DataSet = New DataSet()
        Dim strCurrentTemp As String = "0"

        bWebFeedError1 = False
        bWebFeedError2 = False
        bWebFeedError3 = False

        Dim xsltFile As String = "descriptions.xsl"
        Dim myTransform As New XslCompiledTransform()
        Dim xmlDescription As DataSet = New DataSet()
        Dim strDescriptions As String = "descriptions.txt"
        Dim rowDescription(5) As Object

        Try
            myTransform.Load(xsltFile)
            myTransform.Transform("http://rss.wunderground.com/auto/rss_full/OH/West_Chester.xml?units=english", strDescriptions)
            xmlDescription.ReadXml(strDescriptions)
            For i = 0 To xmlDescription.Tables(0).Rows.Count - 1
                rowDescription(i) = xmlDescription.Tables(0).Rows(i).ItemArray(0)
            Next

            strForecast1 = CleanUp(rowDescription(1).ToString.Trim())
            strForecast2 = CleanUp(rowDescription(2).ToString.Trim())
            strForecast3 = CleanUp(rowDescription(3).ToString.Trim())
        Catch ex As Exception
            bWebFeedError3 = True
            'Exit Sub
        End Try

        ''''''''''

        Try
            Dim xsltFile2 As String = "temps.xsl"
            Dim myTransform2 As New XslCompiledTransform()
            Dim xmlTemps As DataSet = New DataSet()
            Dim strTemps As String = "temps.txt"

            Try
                myTransform2.Load(xsltFile2)
                myTransform2.Transform("http://weather.yahooapis.com/forecastrss?p=45069&u=f", strTemps)
                xmlTemps.ReadXml(strTemps)
            Catch ex As Exception
                'MsgBox("Caught Exception")
                'Exit Sub
            End Try

            'Get Temps
            strCurrentCondition = xmlTemps.Tables(1).Rows(0).ItemArray(0).ToString.Trim & " degrees and " & xmlTemps.Tables(1).Rows(0).ItemArray(1).ToString.Trim
            strCurrentTemp = xmlTemps.Tables(1).Rows(0).ItemArray(1).ToString.Trim
            iTodayHigh = Int(xmlTemps.Tables(2).Rows(0).ItemArray(0).ToString.Trim)
            iTodayLow = Int(xmlTemps.Tables(2).Rows(0).ItemArray(1).ToString.Trim)
            iTomorrowHigh = Int(xmlTemps.Tables(2).Rows(1).ItemArray(0).ToString.Trim)
            iTomorrowLow = Int(xmlTemps.Tables(2).Rows(1).ItemArray(1).ToString.Trim)
        Catch e As Exception
            bWebFeedError1 = True
            rssFeed1 = Nothing
            'Exit Sub
        End Try


        '''''''''''''''

        Try
            rssFeed2 = WebRequest.Create("http://weather.msn.com/RSS.aspx?wealocations=wc:USOH1019&weadegreetype=F")
            rssFeed2.Timeout = 2000
            rssData2.ReadXml(rssFeed2.GetResponse().GetResponseStream())
            rssFeed2 = Nothing
            Dim strPrecipToday, strPrecipTomorrow, strTemp As String
            Dim iParse, iLength As Integer

            Dim chPrecip As Object() = rssData2.Tables(3).Rows(1).ItemArray
            Dim cDescription2 As Integer = rssData2.Tables(3).Columns("description").Ordinal
            Dim strPrecipText As String

            strPrecipText = chPrecip.GetValue(cDescription2).ToString()

            strPrecipToday = ""
            strPrecipTomorrow = ""

            'Get Today's Precip
            iParse = strPrecipText.IndexOf("precipitation: ")
            iParse = iParse + 14
            For iLength = 1 To 3
                iParse = iParse + 1
                strTemp = strPrecipText.Substring(iParse, 1)
                If strTemp <> "%" Then
                    strPrecipToday = strPrecipToday & strTemp
                Else
                    Exit For
                End If
            Next iLength
            iPrecipToday = Int(strPrecipToday)
            strPrecipText = strPrecipText.Remove(0, iParse)

            'Get Tomorrow's Precip
            iParse = strPrecipText.IndexOf("precipitation: ")
            iParse = iParse + 14
            For iLength = 1 To 3
                iParse = iParse + 1
                strTemp = strPrecipText.Substring(iParse, 1)
                If strTemp <> "%" Then
                    strPrecipTomorrow = strPrecipTomorrow & strTemp
                Else
                    Exit For
                End If
            Next iLength
            iPrecipTomorrow = Int(strPrecipTomorrow)

            'Get Current Image
            strPrecipText = rssData2.Tables(3).Rows(0).ItemArray(4).ToString

            iParse = strPrecipText.IndexOf("<img src=")
            iParse = iParse + 50
            strImgCurrent = strPrecipText.Substring(iParse, 2)
            If strImgCurrent(1) = "." Then
                strImgCurrent = strImgCurrent.Substring(0, 1)
            End If
            strImgCurrent = strImgCurrent & ".png"
            PicCurrent.Image = Image.FromFile("C:\program files\watchdog\Images\" & strImgCurrent)
            lblCurrentTemp.Text = strCurrentTemp

            'Get Today's Image

            strPrecipText = rssData2.Tables(3).Rows(1).ItemArray(4).ToString

            iParse = strPrecipText.IndexOf("<img src=")
            iParse = iParse + 50
            strImgToday = strPrecipText.Substring(iParse, 2)
            If strImgToday(1) = "." Then
                strImgToday = strImgToday.Substring(0, 1)
            End If
            strImgToday = strImgToday & ".png"
            PicToday.Image = Image.FromFile("C:\program files\watchdog\Images\" & strImgToday)
            lblHighToday.Text = iTodayHigh.ToString
            lblLowToday.Text = iTodayLow.ToString

            'Get Tomorrow's Image
            strPrecipText = strPrecipText.Substring(iParse, strPrecipText.Length - iParse)

            iParse = strPrecipText.IndexOf("<img src=")
            iParse = iParse + 50
            strImgTomorrow = strPrecipText.Substring(iParse, 2)
            If strImgTomorrow(1) = "." Then
                strImgTomorrow = strImgTomorrow.Substring(0, 1)
            End If
            strImgTomorrow = strImgTomorrow & ".png"
            PicTomorrow.Image = Image.FromFile("C:\program files\watchdog\Images\" & strImgTomorrow)
            lblHighTomorrow.Text = iTomorrowHigh.ToString
            lblLowTomorrow.Text = iTodayLow.ToString

        Catch e As Exception
            bWebFeedError2 = True
            'Exit Sub
        End Try

        '''''''''''

        If Not bWebFeedError1 Then
            'Wear Today
            strWearToday = "Today, you should wear "
            If iTodayHigh > 70 And iTodayLow > 60 Then
                strWearToday = strWearToday & "Short pants and a T Shirt.  "
            ElseIf iTodayHigh > 60 And iTodayLow > 50 Then
                strWearToday = strWearToday & "Long pants and a short sleve shirt with a light jacket or sweater.  "
            ElseIf iTodayHigh > 50 And iTodayLow > 40 Then
                strWearToday = strWearToday & "Long pants and a long sleve shirt with a light jacket or sweater.  "
            ElseIf iTodayHigh <= 50 Or iTodayLow <= 40 Then
                strWearToday = strWearToday & "Long pants and a long sleve shirt with a heavy jacket.  "
                If iTodayHigh <= 35 Then
                    strWearToday = strWearToday & "Bring a Scarf and Mittens.  "
                End If
            End If

            If iTodayHigh > 32 And iPrecipToday >= 40 Then
                strWearToday = strWearToday & "Also, bring an umbrella."
            End If

            'Wear Tomorrow
            strWearTomorrow = "Tomorrow, you should wear "
            If iTomorrowHigh > 70 And iTomorrowLow > 60 Then
                strWearTomorrow = strWearTomorrow & "Short pants and a T Shirt.  "
            ElseIf iTomorrowHigh > 60 And iTomorrowLow > 50 Then
                strWearTomorrow = strWearTomorrow & "Long pants and a short sleve shirt with a light jacket or sweater.  "
            ElseIf iTomorrowHigh > 50 And iTomorrowLow > 40 Then
                strWearTomorrow = strWearTomorrow & "Long pants and a long sleve shirt with a light jacket or sweater.  "
            ElseIf iTomorrowHigh <= 50 Or iTomorrowLow <= 40 Then
                strWearTomorrow = strWearTomorrow & "Long pants and a long sleve shirt with a heavy jacket.  "
                If iTomorrowHigh <= 35 Then
                    strWearTomorrow = strWearTomorrow & "Bring a Scarf and Mittens.  "
                End If
            End If

            If iTomorrowHigh > 32 And iPrecipTomorrow >= 40 Then
                strWearTomorrow = strWearTomorrow & "Also, bring an umbrella."
            End If
        End If

        '''''''''''''''
        If Not bWebFeedError1 And Not bWebFeedError2 Then
            strWaterOff = "No, leave the water on."
            Select Case Month(Now)
                Case 11, 12, 1, 2, 3
                    strWaterOff = "Yes, turn the water off."
                Case 4, 5, 6, 7, 8, 9, 10
                    If iPrecipToday >= 40 Or iPrecipTomorrow >= 40 Or iTodayLow < 40 Or iTomorrowLow < 40 Then
                        strWaterOff = "Yes, turn the water off."
                    End If
            End Select
        End If

    End Sub

    Private Function CleanUp(ByRef strText As String) As String

        strText = strText.Replace("...", ".  ")
        strText = strText.Replace("winds", "wins")
        strText = strText.Replace("Winds", "Wins")
        strText = strText.Replace("wind", "win")
        strText = strText.Replace("Wind", "Win")

        CleanUp = strText

    End Function

#End Region

#Region "Insteon Interface"

    Public Sub Light1_On()
        Insteon_ControlTableAdapter.Request_State_Change(1, "1A.F4.47")
    End Sub

    Public Sub Light1_Off()
        Insteon_ControlTableAdapter.Request_State_Change(0, "1A.F4.47")
    End Sub

    Public Sub RaiseTemp(ByVal iDegrees As Integer)
        UpdateTemp()
        If Temp_Current_StateDataGridView.Rows(0).Cells(7).Value <= 70 Then
            Temp_ControlTableAdapter.Request_Temp_Change("H", "+", 1)
        Else
            Temp_ControlTableAdapter.Request_Temp_Change("C", "+", 1)
        End If

    End Sub

    Public Sub LowerTemp(ByVal iDegrees As Integer)
        UpdateTemp()
        If Temp_Current_StateDataGridView.Rows(0).Cells(7).Value <= 70 Then
            Temp_ControlTableAdapter.Request_Temp_Change("H", "-", 1)
        Else
            Temp_ControlTableAdapter.Request_Temp_Change("C", "-", 1)
        End If
    End Sub

    Public Sub OutsideLightsOn()
        Insteon_ControlTableAdapter.Request_State_Change(1, "15.FD.A1")  'FD Coach
        Thread.Sleep(100)
        Insteon_ControlTableAdapter.Request_State_Change(1, "17.A6.A9")  'GD Coach
        tLightsOff.Start()
    End Sub

    Public Sub OutsideLightsOff()
        tLightsOff.Stop()
        Insteon_ControlTableAdapter.Request_State_Change(0, "15.FD.A1")  'FD Coach
        Thread.Sleep(100)
        Insteon_ControlTableAdapter.Request_State_Change(0, "17.A6.A9")  'GD Coach

    End Sub

#End Region

    Private Sub UpdateHouse()

        Try
            Me.Event_Current_StateTableAdapter.Fill(Me.WatchdogDataSet1.Event_Current_State)
            bDataFeedError = False
        Catch e As Exception
            bDataFeedError = True
        End Try

    End Sub

    Private Sub UpdateTemp()

        Try
            Me.Temp_Current_StateTableAdapter.Fill(Me.WatchdogDataSet1.Temp_Current_State)
            bDataFeedError = False
        Catch e As Exception
            bDataFeedError = True
        End Try

    End Sub

    Private Sub Unmute()
        Try
            lblStatus.Text = "Listening"
            'Event_HistoryTableAdapter.InsertQuery("9007", Now)
            Recognizer.RecognizeAsync(RecognizeMode.Multiple)
            bMute = False
        Catch e As Exception
        End Try

    End Sub

    Private Sub Mute()
        Try
            Recognizer.RecognizeAsyncCancel()
            bMute = True
            'Event_HistoryTableAdapter.InsertQuery("9008", Now)
            lblStatus.Text = "Muted"
        Catch e As Exception

        End Try
    End Sub
#End Region

    Private Sub btnControls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnControls.Click
        If btnControls.Text = "Show" Then
            gbMonitor.Show()
            gbVoiceRecognition.Show()
            btnControls.Text = "Hide"
        Else
            gbMonitor.Hide()
            gbVoiceRecognition.Hide()
            btnControls.Text = "Show"
        End If

    End Sub

#Region "Timers"

    Private Sub tHideWindow_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tHideWindow.Tick

        Try
            iSleepCounter = iSleepCounter - 1
            lblCountdown.Text = iSleepCounter.ToString

            If iSleepCounter <= 0 Then
                tHideWindow.Stop()
                Me.Show()
                bDoorBellPressed = False
                Mute()
                serialPort1.Write("0")
                Start_Monitor()
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub tLightsOff_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tLightsOff.Tick

        OutsideLightsOff()
    End Sub

    Private Sub tSleep_Elapsed()

        Try
            If Not bMute And bMonitoring Then
                iSleepCounter = iSleepCounter - 1
                lblCountdown.Text = iSleepCounter.ToString

                If iSleepCounter <= 0 Then
                    tSleep.Stop()
                    Mute()

                    If Not bDoorBellPressed Then
                        'Turn off monitor
                        serialPort1.Write("0")
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tCheckPresence_Elapsed(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Triggers twice per second
        Dim i As Integer
        Dim BytesRead As Integer
        Dim BufferPointer As Integer
        Dim CBuffer(40) As Char

        Try
            If bMonitoring Then
                If bMute Then
                    BytesRead = serialPort1.Read(CBuffer, BufferPointer, serialPort1.BytesToRead)
                    If BytesRead > 0 Then
                        'search for new command in buffer
                        For j As Short = 0 To (BytesRead - 1)
                            If CBuffer(j) = CChar("1") Or CBuffer(j) = CChar("X") Then
                                iSleepCounter = SLEEPWAIT
                                If CBuffer(j) = CChar("1") Then
                                    serialPort1.Write("1")  'Turn on Monitor
                                End If
                                Synthesizer.Speak("Hello")
                                Unmute()

                                tSleep.SynchronizingObject = Me
                                tSleep.Start()  'Start countdown to go back to sleep
                                serialPort1.DiscardInBuffer()
                            End If
                        Next
                        BufferPointer = 0
                    End If
                Else
                    serialPort1.DiscardInBuffer()
                End If

                'Check doorbell
                i = Event_Current_StateTableAdapter.CheckDoorbell()
                If i = 1 And Not bDoorBellPressed Then
                    Show_Camera()
                End If
            End If
        Catch ex As Exception
            'Do nothing
        End Try

    End Sub
#End Region

End Class
