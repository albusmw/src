Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

  Public Class PlaneWaveEFA

    Private Enum Commands As Byte
      MTR_GET_POS = &H1
      MTR_GOTO_OVER = &H13
      TEMP_GET = &H26
    End Enum

    Private Shared NUM_placeholder As Byte = &H0
    Private Shared CHK_placeholder As Byte = &H0

    Private Shared SOM As Byte = &H3B
    Private Shared SRC_PC As Byte = &H20
    Private Shared RCV_FOC As Byte = &H12

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

    '''<summary>Command to get position.</summary>
    Public Shared Function MTR_GET_POS() As Byte()
      Dim Buffer As New List(Of Byte)
      Buffer.Add(SOM)
      Buffer.Add(NUM_placeholder)
      Buffer.Add(SRC_PC)
      Buffer.Add(RCV_FOC)
      Buffer.Add(Commands.MTR_GET_POS)
      Buffer.Add(CHK_placeholder)
      Dim RetVal As Byte() = Buffer.ToArray
      SetNUM(RetVal)
      SetCHK(RetVal)
      Return RetVal
    End Function

    '''<summary>Command to get motor status.</summary>
    Public Shared Function MTR_GOTO_OVER() As Byte()
      Dim Buffer As New List(Of Byte)
      Buffer.Add(SOM)
      Buffer.Add(NUM_placeholder)
      Buffer.Add(SRC_PC)
      Buffer.Add(RCV_FOC)
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
      Buffer.Add(SRC_PC)
      Buffer.Add(RCV_FOC)
      Buffer.Add(Commands.TEMP_GET)
      Buffer.Add(Sensor)
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
      If Answer.Length >= 8 Then
        Return GetTemperature((Answer(Answer.GetUpperBound(0) - 2) * 256) + Answer(Answer.GetUpperBound(0) - 1))
      Else
        Return Double.NaN
      End If
    End Function

    Public Shared Function GetTemperature(ByVal RawTemp As Integer) As Double
      Dim TempIsNegative As Boolean = False
      If RawTemp > 32768 Then
        TempIsNegative = True
        RawTemp = 65536 - RawTemp
      End If
      Dim IntPart As Integer = RawTemp \ 16
      Dim FractionDigits As Integer = CInt((RawTemp - IntPart) * (625 / 1000))
      Dim RetVal As Double = IntPart + (FractionDigits / 10)
      If TempIsNegative Then RetVal = -RetVal
      Return RetVal
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