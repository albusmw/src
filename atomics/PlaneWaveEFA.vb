Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Class to communicate with the planewave equipment (EFA, Delta-T).</summary>
    '''<seealso cref="C:\Users\albus\Dropbox\Astro\Unterlagen Ausrüstung\PlaneWave\EFA-Communication-Protocols.pdf"/>
    '''<seealso cref="C:\Users\albus\Dropbox\Astro\Unterlagen Ausrüstung\PlaneWave\EFA-Communication-Protocols 2.pdf"/>
    '''<seealso cref="C:\Users\albus\Dropbox\Astro\Unterlagen Ausrüstung\PlaneWave\Delta-T Communication Protocol.pdf"/>
    '''<seealso cref="C:\Users\albus\Dropbox\Astro\Unterlagen Ausrüstung\PlaneWave\efalib_for_customers.txt"/>
    '''<remarks>Contact at PlaneWave: Kevin Ivarsen <kivarsen@gmail.com> / Jason Fournier <jfournier@planewave.com></remarks>
    Public Class cPWI_IO

        Public Event Log(ByVal Message As String)

        Public Event LogCOMIO(ByVal Message As String)

        Public Enum eOnOff
            Unknown = -1
            Off = 0
            [On] = 1
        End Enum

        '''<summary>Available commands.</summary>
        Private Enum Commands As Byte
            '''<summary>EFA: Get position.</summary>
            MTR_GET_POS = &H1
            '''<summary>EFA: Determine if the motor is moving during a GOTO?</summary>
            MTR_GOTO_OVER = &H13
            '''<summary>EFA: Get the temperature of one sensor.</summary>
            TEMP_GET = &H26
            '''<summary>EFA: Set the fans, on or off.</summary>
            FANS_SET = &H27
            '''<summary>EFA: Get the fans state, on or off.</summary>
            FANS_GET = &H28
            '''<summary>EFA: Move the motor positive. Motor will stop when the Max Slew Limit is reached.</summary>
            MTR_PMSLEW_RATE = &H24
            '''<summary>EFA: Move the motor negative. Motor will stop when the Max Slew Limit is reached.</summary>
            MTR_NMSLEW_RATE = &H25
            '''<summary>Delta-T: Returns the number of available heaters that can be indexed.</summary>
            COH_NUMHEATERS = &HB0
            '''<summary>Delta-T: Returns a packed structure with data on the current operational state for the specified heater.</summary>
            COH_REPORT = &HB5
            '''<summary>Delta-T: Rescans 1-Wire bus and updates temp sensors.</summary>
            COH_RESCAN = &HBF
        End Enum

        '''<summary>Available commands.</summary>
        Private Enum MsgSpecific As Byte
            '''<summary>Placeholder for number of bytes - to be replaced with real number calculated.</summary>
            NUM_placeholder = &H0
            '''<summary>Placeholder for CRC - to be replaced with real checksum calculated.</summary>
            CHK_placeholder = &H0
            '''<summary>Start Of Message byte - must start every message packet.</summary>
            SOM = &H3B
        End Enum

        '''<summary>Available commands.</summary>
        Public Enum eAddr As Byte
            '''<summary>Address of PC.</summary>
            PC = &H20
            '''<summary>Address of Focuser.</summary>
            Focuser = &H12
            '''<summary>Address of fan controller.</summary>
            FanController = &H13
            '''<summary>Address of temperature sensor - same as focuser ...</summary>
            TemperatureSensor = &H12
            '''<summary>Address of temperature sensor - same as focuser ...</summary>
            DeltaT = &H32
        End Enum

        Public Function GetPosition(ByRef Port As IO.Ports.SerialPort) As Integer

            RaiseEvent Log("Running command <Get position> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = MTR_GET_POS()

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            '-----------------------------------------------------
            'Decode
            Return MTR_GET_POS_decode(AnswerBuffer)

        End Function

        Public Function GetGotoOver(ByRef Port As IO.Ports.SerialPort) As Boolean

            RaiseEvent Log("Running command <Get GOTO over (Goto finished?)> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = MTR_GOTO_OVER()

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            '-----------------------------------------------------
            'Decode
            Return MTR_GOTO_OVER_decode(AnswerBuffer)

        End Function

        Public Function GetEFATemp(ByRef Port As IO.Ports.SerialPort, ByVal Target As eAddr, ByVal Sensor As Byte, ByRef Status As String) As Double

            RaiseEvent Log("Running command <Get T of sensor " & Sensor.ToString.Trim & "> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = TEMP_GET(Target, Sensor)

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            '-----------------------------------------------------
            'Decode
            Dim RetVal As Double = TEMP_GET_decode(AnswerBuffer, Status)
            RaiseEvent Log("T[" & Target.ToString.Trim & ":" & Sensor.ToString.Trim & "            : <" & RetVal & "> (" & Status & ")")
            Return RetVal

        End Function

        Public Function GetFans(ByRef Port As IO.Ports.SerialPort) As Ato.cPWI_IO.eOnOff

            RaiseEvent Log("Running command <Get Fans status> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = FANS_GET()

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            '-----------------------------------------------------
            'Decode
            Dim RetVal As Ato.cPWI_IO.eOnOff = FANS_GET_decode(AnswerBuffer)
            RaiseEvent Log("Fans            : <" & RetVal & ">")
            Return RetVal

        End Function

        Public Function SetFans(ByRef Port As IO.Ports.SerialPort, ByVal State As Boolean) As Boolean

            RaiseEvent Log("Running command <Set Fans status " & CStr(State) & "> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = FANS_SET(State)

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            Return True

        End Function

        Public Function MoveFocuser(ByRef Port As IO.Ports.SerialPort, ByVal Positive As Boolean, ByVal Speed As Byte) As Boolean

            RaiseEvent Log("Running command <Move focuser " & CStr(IIf(Positive, "+", "-")) & ", speed " & CStr(Speed) & "> ...")
            PrepareCOM(Port)
            Dim CommandBuffer As Byte() = MTR_SLEW_RATE(Positive, Speed)

            Dim AnswerBuffer As Byte() = {}
            RaiseEvent LogCOMIO(" -> Status: " & EFACommunication(Port, CommandBuffer, AnswerBuffer))

            Return True

        End Function

        Public Sub PrepareCOM(ByRef Port As IO.Ports.SerialPort)
            '-----------------------------------------------------
            'Read remaining bytes
            If Port.BytesToRead > 0 Then
                Dim DummyBuffer(Port.BytesToRead - 1) As Byte : Port.Read(DummyBuffer, 0, Port.BytesToRead)
            End If
            '-----------------------------------------------------
            'Wait for the CTS to be clear
            WaitForCTS(Port, False)
            '-----------------------------------------------------
            'Try to own bus
            Port.RtsEnable = True
            WaitForCTS(Port, True)
        End Sub

        Public Function ValidateEcho(ByRef Port As IO.Ports.SerialPort, ByRef CommandBuffer As Byte()) As Boolean

            '-----------------------------------------------------
            'Wait for echo
            Do
                System.Threading.Thread.Sleep(10)
            Loop Until Port.BytesToRead >= CommandBuffer.Length

            '-----------------------------------------------------
            'Read echo
            Dim EchoBuffer(CommandBuffer.Length - 1) As Byte
            Port.Read(EchoBuffer, 0, EchoBuffer.Length)

            '-----------------------------------------------------
            'RTS to low
            Port.RtsEnable = False

            '-----------------------------------------------------
            'Validate echo
            For Idx As Integer = 0 To CommandBuffer.GetUpperBound(0)
                If EchoBuffer(Idx) <> CommandBuffer(Idx) Then Return False
            Next Idx

            Return True

        End Function

        Public Function ReadAnswer(ByRef Port As IO.Ports.SerialPort) As Byte()

            '-----------------------------------------------------
            'Wait for avaiable bytes
            Do
                System.Threading.Thread.Sleep(10)
            Loop Until Port.BytesToRead >= 1
            System.Threading.Thread.Sleep(10)

            '-----------------------------------------------------
            'Read
            Dim AnswerBuffer(Port.BytesToRead - 1) As Byte
            Port.Read(AnswerBuffer, 0, AnswerBuffer.Length)

            '-----------------------------------------------------
            'Return
            Return AnswerBuffer

        End Function

        '''<summary>Command to get focuser position.</summary>
        Public Function MTR_GET_POS() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(eAddr.Focuser)
            Buffer.Add(Commands.MTR_GET_POS)
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to get focuser motor status.</summary>
        Public Function MTR_GOTO_OVER() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(eAddr.Focuser)
            Buffer.Add(Commands.MTR_GOTO_OVER)
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to get temperature value.</summary>
        '''<param name="Sensor">Sensor to query - Primary=0, Ambient=1, Secondary=2.</param>
        Public Function TEMP_GET(ByVal Target As eAddr, ByVal Sensor As Byte) As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(Target)
            Buffer.Add(Commands.TEMP_GET)
            Buffer.Add(Sensor)
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to get fans state.</summary>
        Public Function FANS_GET() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(eAddr.FanController)
            Buffer.Add(Commands.FANS_GET)
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to set fans state.</summary>
        '''<param name="State">TRUE for ON, FALSE for OFF.</param>
        Public Function FANS_SET(ByVal State As Boolean) As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(eAddr.FanController)
            Buffer.Add(Commands.FANS_SET)
            Buffer.Add(CByte(IIf(State = True, 1, 0)))
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to move the focuser positive or negative.</summary>
        '''<param name="Positive">TRUE for positive movement, FALSE for negative.</param>
        '''<param name="Speed">Movement speed - 0 (stop) to 9 (maximum).</param>
        Public Function MTR_SLEW_RATE(ByVal Positive As Boolean, ByVal Speed As Byte) As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(MsgSpecific.SOM)
            Buffer.Add(MsgSpecific.NUM_placeholder)
            Buffer.Add(eAddr.PC)
            Buffer.Add(eAddr.Focuser)
            If Positive Then
                Buffer.Add(Commands.MTR_PMSLEW_RATE)
            Else
                Buffer.Add(Commands.MTR_NMSLEW_RATE)
            End If
            Buffer.Add(Speed)
            Buffer.Add(MsgSpecific.CHK_placeholder)
            Return FormatMessage(Buffer)
        End Function

        '''<summary>Command to parse the answer of position.</summary>
        Public Function MTR_GET_POS_decode(ByRef Answer As Byte()) As Integer
            If Answer.Length >= 9 Then
                Return ((Answer(Answer.GetUpperBound(0) - 3)) * (256 * 256)) + ((Answer(Answer.GetUpperBound(0) - 2)) * 256) + (Answer(Answer.GetUpperBound(0) - 1))
            Else
                Return 0
            End If
        End Function

        '''<summary>Command to parse the answer of GOTO over.</summary>
        Public Function MTR_GOTO_OVER_decode(ByRef Answer As Byte()) As Boolean
            If Answer.Length >= 7 Then
                Return CBool(IIf(Answer(Answer.GetUpperBound(0) - 1) = 255, True, False))
            Else
                Return False
            End If
        End Function

        '''<summary>Command to parse the answer of TEMP_GET over.</summary>
        Public Function TEMP_GET_decode(ByRef Answer As Byte(), ByRef Status As String) As Double
            Dim RetVal As Integer = Integer.MinValue
            Dim NoSensor As Integer = &H7F7F
            Status = String.Empty
            If Answer.Length >= 8 Then
                RetVal = Answer(Answer.GetUpperBound(0) - 1) * 256 + (Answer(Answer.GetUpperBound(0) - 2))
                If RetVal = NoSensor Then
                    Status = "No sensor for the address requested"
                    Return -1000
                End If
                If RetVal > &H8000 Then RetVal = RetVal - &H10000
            Else
                Status = "Wrong answer length"
                Return Double.NaN
            End If
            Return RetVal * (1 / 16)
        End Function

        '''<summary>Command to parse the answer of FANS_GET.</summary>
        Public Function FANS_GET_decode(ByRef Answer As Byte()) As eOnOff
            If Answer.Length >= 7 Then
                Return CType(IIf(Answer(Answer.GetUpperBound(0) - 1) = 3, eOnOff.Off, eOnOff.On), eOnOff)
            Else
                Return Ato.cPWI_IO.eOnOff.Unknown
            End If
        End Function

        '''<summary>Command to parse the answer of position.</summary>
        Public Function GetPositionMicrons(ByVal Position As Integer) As Double
            Return 1000 * (Position / 115134.42)
        End Function

        '''<summary>Compose a hex string for the byte vector (for debug only).</summary>
        Public Function AsHex(ByRef Vector As Byte()) As String
            If IsNothing(Vector) = True Then Return String.Empty
            Dim RetVal As String = String.Empty
            For Each Entry As Byte In Vector
                RetVal &= "0x" & CStr(IIf(Hex(Entry).Length = 1, "0" & Hex(Entry), Hex(Entry))) & " "
            Next Entry
            Return RetVal.Trim
        End Function

        Public Function FormatMessage(ByRef Buffer As List(Of Byte)) As Byte()
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Set the NUM byte in the command string.</summary>
        '''<remarks>Number of Bytes is calculated by: NUM = [(Packet Byte Count) - 3] (Note: The bytes SOM, NUM, and CHK are not counted in NUM)</remarks>
        Private Sub SetNUM(ByRef Bytes() As Byte)
            Bytes(1) = CByte(Bytes.Length - 3)
        End Sub

        '''<summary>Set the Checksum byte in the command string.</summary>
        Private Sub SetCHK(ByRef Bytes() As Byte)
            Dim Sum As Integer = 0
            For Idx As Integer = 1 To Bytes.GetUpperBound(0) - 1
                Sum += Bytes(Idx)
            Next Idx
            Dim IntBytes() As Byte = BitConverter.GetBytes(Sum)
            Bytes(Bytes.GetUpperBound(0)) = TwosComplement(IntBytes(0))
        End Sub

        Private Function TwosComplement(value As Byte) As Byte
            If value = 0 Then Return 0 Else Return CByte(CByte(value Xor Byte.MaxValue) + 1)
        End Function

        Private Sub WaitForCTS(ByRef Port As IO.Ports.SerialPort, ByVal RequiredState As Boolean)
            If Port.CtsHolding <> RequiredState Then
                Do
                    System.Threading.Thread.Sleep(10)
                Loop Until Port.CtsHolding = RequiredState
            End If
        End Sub

        Private Function EFACommunication(ByRef EFAPort As IO.Ports.SerialPort, ByRef CommandBuffer As Byte(), ByRef AnswerBuffer As Byte()) As Boolean
            LogWrite(CommandBuffer)
            EFAPort.Write(CommandBuffer, 0, CommandBuffer.Length)
            Dim RetVal As Boolean = ValidateEcho(EFAPort, CommandBuffer)
            AnswerBuffer = ReadAnswer(EFAPort)
            LogRead(AnswerBuffer)
            Return RetVal
        End Function

        Private Sub LogWrite(ByRef Buffer As Byte())
            If IsNothing(Buffer) = True Then
                RaiseEvent LogCOMIO(">> 0 byte")
            Else
                RaiseEvent LogCOMIO(">> " & AsHex(Buffer) & " (" & Buffer.Length.ToString.Trim & " byte)")
            End If
        End Sub

        Private Sub LogRead(ByRef Buffer As Byte())
            If IsNothing(Buffer) = True Then
                RaiseEvent LogCOMIO("<< 0 byte")
            Else
                RaiseEvent LogCOMIO("<< " & AsHex(Buffer) & " (" & Buffer.Length.ToString.Trim & " byte)")
            End If
        End Sub

    End Class

End Namespace