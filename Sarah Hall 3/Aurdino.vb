Public Class Arduino_Net2
    Private WithEvents serialPort1 As System.IO.Ports.SerialPort
    Private CommandBuffer As String
    Private Receiving As Boolean
    Private BufferPointer As Integer
    Private CBuffer(40) As Byte

    Private _ComPort As String = "COM8"
    Private _BaudRate As Integer = 9600

    'Private WithEvents DataCheckTimer As System.Timers.Timer
    Private WithEvents WatchDogTimer As System.Timers.Timer

    Public Enum DigitalDirection
        Input = 0
        DigitalOutput = 1
        PWMOutput = 2
    End Enum


    Public Event DigitalDataReceived(ByVal DPortNr As Integer, ByVal Value As Integer)

    Public Event AnalogDataReceived(ByVal APortNr As Integer, ByVal Value As Integer)

    Public Event LogMessageReceived(ByVal Message As String, ByVal DebugLevel As Integer)

    Public Event WatchdogReceived()

    Public Event ConnectionLost()

    Public Property ComPort() As String
        Get
            Return _ComPort
        End Get
        Set(ByVal value As String)
            _ComPort = value
        End Set
    End Property

    Public Property BaudRate() As String
        Get
            Return _BaudRate
        End Get
        Set(ByVal value As String)
            _BaudRate = value
        End Set
    End Property

    Public Sub New(ByVal PortName As String)
        Me.ComPort = PortName.Trim
        'DataCheckTimer = New System.Timers.Timer(50)
        'AddHandler DataCheckTimer.Elapsed, AddressOf CheckTimerElapsed
    End Sub

    Public Function StartCommunication() As Boolean
        Dim _ReturnValue As Boolean = False
        Try
            Dim components As System.ComponentModel.IContainer = New System.ComponentModel.Container()
            serialPort1 = New System.IO.Ports.SerialPort(components)
            serialPort1.PortName = _ComPort
            serialPort1.BaudRate = _BaudRate
            serialPort1.ReceivedBytesThreshold = 1

            serialPort1.Open()
            If Not serialPort1.IsOpen Then
                WriteDebug("Unable to open comport...")
                _ReturnValue = False
                Exit Function
            Else
                serialPort1.DtrEnable = True
                WriteDebug("Serial port is open")
                System.Threading.Thread.Sleep(1000)
                Dim Command As Byte() = {40, 0, 0, 0, 41, 0}
                'send some empty commands to clear all buffers on Arduino
                Me.SendCommand(Command)
                Me.SendCommand(Command)
                'DataCheckTimer.Enabled = True
                'DataCheckTimer.Start()
                WatchDogTimer = New System.Timers.Timer(5000)
                WatchDogTimer.Enabled = True
                WatchDogTimer.Start()
                _ReturnValue = True
            End If

            ' callback for text coming back from the arduino
            AddHandler serialPort1.DataReceived, AddressOf OnReceived
        Catch ex As Exception
            WriteDebug("Error opening comport...")
            _ReturnValue = False
        End Try
        Return _ReturnValue
    End Function

    Public Sub SetDigitalDirection(ByVal Port As Integer, ByVal Direction As DigitalDirection)
        'TODO check value and portnumber
        Dim Command1 As Byte() = {40, 82, Port, Direction, 41, 0}
        Me.SendCommand(Command1)
    End Sub

    Public Sub SetDigitalValue(ByVal Port As Integer, ByVal Value As Integer)
        'TODO check value and portnumber
        Dim Command1 As Byte() = {40, 80, Port, Value, 41, 0}
        Me.SendCommand(Command1)
    End Sub

    Public Sub EnableDigitalPort(ByVal Port As Integer, ByVal Enable As Boolean)
        'TODO check value and portnumber
        If Enable Then
            Dim Command1 As Byte() = {40, 68, Port, 1, 41, 0}
            Me.SendCommand(Command1)
        Else
            Dim Command1 As Byte() = {40, 68, Port, 0, 41, 0}
            Me.SendCommand(Command1)
        End If
    End Sub

    Public Sub EnableAnalogPort(ByVal Port As Integer, ByVal Enable As Boolean)
        'TODO check value and portnumber
        If Enable Then
            Dim Command1 As Byte() = {40, 65, Port, 1, 41, 0}
            Me.SendCommand(Command1)
        Else
            Dim Command1 As Byte() = {40, 65, Port, 0, 41, 0}
            Me.SendCommand(Command1)
        End If
    End Sub

    Public Sub EnableDigitalTrigger(ByVal Port As Integer, ByVal Enable As Boolean)
        'TODO check value and portnumber
        If Enable Then
            Dim Command1 As Byte() = {40, 84, Port, 1, 41, 0}
            Me.SendCommand(Command1)
        Else
            Dim Command1 As Byte() = {40, 84, Port, 0, 41, 0}
            Me.SendCommand(Command1)
        End If
    End Sub

    Public Sub EnableAnalogTrigger(ByVal Port As Integer, ByVal Threshold As Integer)
        'TODO check value and portnumber
        Dim Command1 As Byte() = {40, 83, Port, Threshold, 41, 0}
        Me.SendCommand(Command1)
    End Sub

    Public Sub GetDigitalValue(ByVal Port As Integer)
        'TODO check value and portnumber
        Dim Command1 As Byte() = {40, 86, Port, 0, 41, 0}
        Me.SendCommand(Command1)
    End Sub

    Public Sub GetAnalogValue(ByVal Port As Integer)
        'TODO check value and portnumber
        Dim Command1 As Byte() = {40, 87, Port, 0, 41, 0}
        Me.SendCommand(Command1)
    End Sub

    Private Sub SendCommand(ByVal Command As Byte())
        Try
            If Command.Length > 0 Then
                If serialPort1.IsOpen Then
                    serialPort1.Write(Command, 0, Command.Length - 1)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub OnReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs)
        If Receiving = False Then
            Receiving = True
            ProcessSerialData()
            Receiving = False
        End If
    End Sub

    Private Sub CheckTimerElapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        If Receiving = False Then
            Receiving = True
            ProcessSerialData()
            Receiving = False
        End If
    End Sub

    Private Sub ProcessSerialData()
        Try
            Dim BytesRead As Integer
            While serialPort1.BytesToRead > 0
                BytesRead = serialPort1.Read(CBuffer, BufferPointer, serialPort1.BytesToRead)
                BufferPointer += BytesRead
                While BufferPointer > 0
                    'search start of new command in buffer
                    Dim CommandStart As Integer = -1
                    For i As Short = 0 To BufferPointer
                        If CBuffer(i) = CByte(40) Then
                            CommandStart = i
                            Exit For
                        End If
                    Next
                    If CommandStart = -1 Then
                        'no begin of command found. this is no valid situation
                        ClearCBuffer()
                    End If
                    If CommandStart > 0 Then
                        'begin of command found somewhere in buffer. dismiss al bytes before startcharacter
                        LeftShiftCBuffer(CommandStart)
                    End If

                    'at this point we should have a clean commandbuffer

                    'search end of command in buffer
                    Dim Commandend As Integer = 0
                    For i As Short = 0 To BufferPointer
                        If CBuffer(i) = CByte(41) Then
                            Commandend = i
                            Exit For
                        End If
                    Next
                    If Commandend > 0 Then
                        'commandend found. there should be a complete command in the buffer
                        Dim CommandBytes(Commandend) As Byte
                        For i As Integer = 0 To Commandend
                            CommandBytes(i) = CBuffer(i)
                        Next
                        'execute command
                        ProcessCommand(CommandBytes)
                        'reset pointer
                        LeftShiftCBuffer(Commandend + 1)
                    Else
                        Exit While
                    End If
                End While
            End While
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ClearCBuffer()
        For i As Integer = 0 To CBuffer.Length - 1
            CBuffer(i) = 0
        Next
        BufferPointer = 0
    End Sub

    Private Sub LeftShiftCBuffer(ByVal NrOfPlaces As Integer)
        For i As Integer = NrOfPlaces To CBuffer.Length - 1
            CBuffer(i - NrOfPlaces) = CBuffer(i)
        Next
        For i As Integer = CBuffer.Length - NrOfPlaces To CBuffer.Length - 1
            CBuffer(i) = 0
        Next
        BufferPointer -= NrOfPlaces
    End Sub

    Private Sub ProcessCommand(ByVal CommandBytes As Byte())
        Try
            If ((CommandBytes(0) = CByte(40)) And (CommandBytes(CommandBytes.Length - 1) = CByte(41))) Then
                Dim PType As Char = ChrW(CommandBytes(1))
                Dim PNumber As Integer
                Dim Value As Integer
                Dim CFound As Boolean = False
                Select Case PType
                    Case "D"
                        PNumber = CInt(CommandBytes(2))
                        Value = CInt(CommandBytes(3))
                        RaiseEvent DigitalDataReceived(PNumber, Value)
                    Case "A"
                        PNumber = CInt(CommandBytes(2))
                        Value = (CInt(CommandBytes(3)) * 255) + CInt(CommandBytes(4))
                        RaiseEvent AnalogDataReceived(PNumber, Value)
                    Case "H"
                        PNumber = 0
                        Value = 0
                        RaiseEvent WatchdogReceived()
                        If Not IsNothing(WatchDogTimer) Then
                            WatchDogTimer.Stop()
                            WatchDogTimer.Dispose()
                        End If
                        WatchDogTimer = New System.Timers.Timer(5000)
                        WatchDogTimer.Enabled = True
                        WatchDogTimer.Start()
                    Case "Q"
                        Dim CommandString As String = String.Empty
                        For i As Integer = 0 To CommandBytes.Length - 1
                            CommandString += CommandBytes(i).ToString + " "
                        Next
                        WriteDebug("Arduino Q msg: " + CommandString)
                    Case Else
                        Dim CommandString As String = String.Empty
                        For i As Integer = 0 To CommandBytes.Length - 1
                            CommandString += CommandBytes(i).ToString + " "
                        Next
                        WriteDebug("Arduino msg: " + CommandString)
                End Select
            Else
                Dim CommandString As String = String.Empty
                For i As Integer = 1 To CommandBytes.Length - 1
                    CommandString += CommandBytes(i).ToString
                Next
                WriteDebug("Error:bad commandformat received: " + CommandString)
            End If
        Catch ex As Exception
            WriteDebug("Error:bad commandformat received")
        End Try
    End Sub

    Private Sub WatchdogTimerElaped(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles WatchDogTimer.Elapsed
        RaiseEvent ConnectionLost()
        If Not IsNothing(WatchDogTimer) Then
            WatchDogTimer.Stop()
            WatchDogTimer.Dispose()
        End If
    End Sub

    Private Sub WriteDebug(ByVal Message As String, Optional ByVal Level As Integer = 0)
        RaiseEvent LogMessageReceived(Message, Level)
        Console.WriteLine(Message & " " & Level.ToString)
    End Sub
End Class
