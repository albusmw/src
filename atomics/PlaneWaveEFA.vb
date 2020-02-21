Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Class to communicate with the planewave equipment (EFA, Delta-T).</summary>
    '''<seealso cref="EFA-Communication-Protocols.pdf"/>
    '''<seealso cref="EFA-Communication-Protocols 2.pdf"/>
    '''<seealso cref="efalib_for_customers.txt"/>
    Public Class PlaneWaveEFA

        '''<summary>Available commands.</summary>
        Private Enum Commands As Byte
            '''<summary>Get position.</summary>
            MTR_GET_POS = &H1
            '''<summary>Determine if the motor is moving during a GOTO?</summary>
            MTR_GOTO_OVER = &H13
            '''<summary>Get the temperature of one sensor.</summary>
            TEMP_GET = &H26
            '''<summary>Set the fans, on or off.</summary>
            FANS_SET = &H27
            '''<summary>Get the fans state, on or off.</summary>
            FANS_GET = &H28
        End Enum

        '''<summary>Placeholder for number of bytes - to be replaced with real number calculated.</summary>
        Private Shared NUM_placeholder As Byte = &H0

        '''<summary>Placeholder for CRC - to be replaced with real checksum calculated.</summary>
        Private Shared CHK_placeholder As Byte = &H0

        '''<summary>Start Of Message byte - must start every message packet.</summary>
        Private Shared SOM As Byte = &H3B

        '''<summary>Address of PC.</summary>
        Private Shared PC As Byte = &H20
        '''<summary>Address of Focuser.</summary>
        Private Shared Focuser As Byte = &H12
        '''<summary>Address of fan controller.</summary>
        Private Shared FanController As Byte = &H13
        '''<summary>Address of temperature sensor.</summary>
        Private Shared TemperatureSensor As Byte = &H12

        Public Shared Sub PrepareCOM(ByRef Port As IO.Ports.SerialPort)
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

        Public Shared Function ValidateEcho(ByRef Port As IO.Ports.SerialPort, ByRef CommandBuffer As Byte()) As Boolean

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

        Public Shared Function ReadAnswer(ByRef Port As IO.Ports.SerialPort) As Byte()

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
        Public Shared Function MTR_GET_POS() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(SOM)
            Buffer.Add(NUM_placeholder)
            Buffer.Add(PC)
            Buffer.Add(Focuser)
            Buffer.Add(Commands.MTR_GET_POS)
            Buffer.Add(CHK_placeholder)
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Command to get focuser motor status.</summary>
        Public Shared Function MTR_GOTO_OVER() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(SOM)
            Buffer.Add(NUM_placeholder)
            Buffer.Add(PC)
            Buffer.Add(Focuser)
            Buffer.Add(Commands.MTR_GOTO_OVER)
            Buffer.Add(CHK_placeholder)
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Command to get temperature value.</summary>
        Public Shared Function TEMP_GET(ByVal Sensor As Byte) As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(SOM)
            Buffer.Add(NUM_placeholder)
            Buffer.Add(PC)
            Buffer.Add(Focuser)
            Buffer.Add(Commands.TEMP_GET)
            Buffer.Add(Sensor)
            Buffer.Add(CHK_placeholder)
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Command to get fans state.</summary>
        Public Shared Function FANS_GET() As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(SOM)
            Buffer.Add(NUM_placeholder)
            Buffer.Add(PC)
            Buffer.Add(FanController)
            Buffer.Add(Commands.FANS_GET)
            Buffer.Add(CHK_placeholder)
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Command to set fans state.</summary>
        Public Shared Function FANS_SET(ByVal State As Boolean) As Byte()
            Dim Buffer As New List(Of Byte)
            Buffer.Add(SOM)
            Buffer.Add(NUM_placeholder)
            Buffer.Add(PC)
            Buffer.Add(FanController)
            Buffer.Add(Commands.FANS_SET)
            Buffer.Add(CByte(IIf(State = True, 1, 0)))
            Buffer.Add(CHK_placeholder)
            Dim RetVal As Byte() = Buffer.ToArray
            SetNUM(RetVal)
            SetCHK(RetVal)
            Return RetVal
        End Function

        '''<summary>Command to parse the answer of position.</summary>
        Public Shared Function MTR_GET_POS_decode(ByRef Answer As Byte()) As Integer
            If Answer.Length >= 9 Then
                Return ((Answer(Answer.GetUpperBound(0) - 3)) * (256 * 256)) + ((Answer(Answer.GetUpperBound(0) - 2)) * 256) + (Answer(Answer.GetUpperBound(0) - 1))
            Else
                Return 0
            End If
        End Function

        '''<summary>Command to parse the answer of GOTO over.</summary>
        Public Shared Function MTR_GOTO_OVER_decode(ByRef Answer As Byte()) As Boolean
            If Answer.Length >= 7 Then
                Return CBool(IIf(Answer(Answer.GetUpperBound(0) - 1) = 255, True, False))
            Else
                Return False
            End If
        End Function

        '''<summary>Command to parse the answer of TEMP_GET over.</summary>
        Public Shared Function TEMP_GET_decode(ByRef Answer As Byte()) As Double
            Dim RetVal As Integer = Integer.MinValue
            If Answer.Length >= 8 Then
                RetVal = Answer(Answer.GetUpperBound(0) - 1) * 256 + (Answer(Answer.GetUpperBound(0) - 2))
                If RetVal = &H7F7F Then Return Double.NaN
                If RetVal > &H8000 Then RetVal = RetVal - &H10000
            Else
                Return Double.NaN
            End If
            Return RetVal * (1 / 16)
        End Function

        '''<summary>Command to parse the answer of FANS_GET.</summary>
        Public Shared Function FANS_GET_decode(ByRef Answer As Byte()) As Boolean
            If Answer.Length >= 7 Then
                Return CBool(IIf(Answer(Answer.GetUpperBound(0) - 1) = 3, False, True))
            Else
                Return False
            End If
        End Function

        '''<summary>Command to parse the answer of position.</summary>
        Public Shared Function GetPositionMicrons(ByVal Position As Integer) As Double
            Return 1000 * (Position / 115134.42)
        End Function

        '''<summary>Compose a hex string for the byte vector (for debug only).</summary>
        Public Shared Function AsHex(ByRef Vector As Byte()) As String
            If IsNothing(Vector) = True Then Return String.Empty
            Dim RetVal As String = String.Empty
            For Each Entry As Byte In Vector
                RetVal &= "0x" & CStr(IIf(Hex(Entry).Length = 1, "0" & Hex(Entry), Hex(Entry))) & " "
            Next Entry
            Return RetVal.Trim
        End Function

        '''<summary>Set the NUM byte in the command string.</summary>
        '''<remarks>Number of Bytes is calculated by: NUM = [(Packet Byte Count) - 3] (Note: The bytes SOM, NUM, and CHK are not counted in NUM)</remarks>
        Private Shared Sub SetNUM(ByRef Bytes() As Byte)
            Bytes(1) = CByte(Bytes.Length - 3)
        End Sub

        '''<summary>Set the Checksum byte in the command string.</summary>
        Private Shared Sub SetCHK(ByRef Bytes() As Byte)
            Dim Sum As Integer = 0
            For Idx As Integer = 1 To Bytes.GetUpperBound(0) - 1
                Sum += Bytes(Idx)
            Next Idx
            Dim IntBytes() As Byte = BitConverter.GetBytes(Sum)
            Bytes(Bytes.GetUpperBound(0)) = TwosComplement(IntBytes(0))
        End Sub

        Private Shared Function TwosComplement(value As Byte) As Byte
            If value = 0 Then Return 0 Else Return CByte(CByte(value Xor Byte.MaxValue) + 1)
        End Function

        Private Shared Sub WaitForCTS(ByRef Port As IO.Ports.SerialPort, ByVal RequiredState As Boolean)
            If Port.CtsHolding <> RequiredState Then
                Do
                    System.Threading.Thread.Sleep(10)
                Loop Until Port.CtsHolding = RequiredState
            End If
        End Sub

    End Class

End Namespace