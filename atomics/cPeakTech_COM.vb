Option Explicit On
Option Strict On

'''<summary>Class for direct PeakTech access.</summary>
Public Class cPeakTech_COM

    Dim COM_port As IO.Ports.SerialPort
    Dim ReadBuffer As String = String.Empty
    Dim StopChars As String = vbCr

    Const CAT_Main As String = "1. Readings"

    '''<summary>Supported PeakTech commands.</summary>
    Private Class Commands
        '''<summary>Maximum voltage and current value.</summary>
        Public Shared ReadOnly Property [GMAX] As String = "GMAX"
        '''<summary>Current voltage, current and output state value (as displayed).</summary>
        Public Shared ReadOnly Property [GETD] As String = "GETD"
    End Class

    '''<summary>Current output voltage [V].</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("1.1. Current voltage [V]")>
    Public ReadOnly Property U() As Double
        Get
            Return MyU
        End Get
    End Property
    Private MyU As Double = Double.NaN

    '''<summary>Current output current [A].</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("1.2. Current current [A]")>
    Public ReadOnly Property I() As Double
        Get
            Return MyI
        End Get
    End Property
    Private MyI As Double = Double.NaN

    '''<summary>Current output power [W].</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("1.3. Current power [W]")>
    Public ReadOnly Property P() As Double
        Get
            Return I * U
        End Get
    End Property

    '''<summary>CC or CV mode.</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("1.4. Output state")>
    Public ReadOnly Property Output() As String
        Get
            Return MyOutput
        End Get
    End Property
    Private MyOutput As String = "???"

    '''<summary>CC or CV mode.</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("1.5. CV or CC mode")>
    Public ReadOnly Property CVCCMode() As String
        Get
            Return MyCVCCMode
        End Get
    End Property
    Private MyCVCCMode As String = String.Empty

    '''<summary>Maximum output voltage [V].</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("2.1. Maximum voltage [V]")>
    Public ReadOnly Property U_max() As Double
        Get
            Return MyU_max
        End Get
    End Property
    Private MyU_max As Double = Double.NaN

    '''<summary>Maximum output current [A].</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("2.2. Maximum current [A]")>
    Public ReadOnly Property I_max() As Double
        Get
            Return MyI_max
        End Get
    End Property
    Private MyI_max As Double = Double.NaN

    '''<summary>Last update.</summary>
    <System.ComponentModel.Category(CAT_Main)>
    <System.ComponentModel.DisplayName("3.1. Last update")>
    Public ReadOnly Property LastUpdate() As DateTime
        Get
            Return MyLastUpdate
        End Get
    End Property
    Private MyLastUpdate As DateTime = Nothing


    Public Sub Init(ByVal SerialPort As String)

        Dim InitCOMNow As Boolean = False
        If IsNothing(COM_port) = True Then
            COM_port = New IO.Ports.SerialPort
            InitCOMNow = True
        Else
            If COM_port.IsOpen = False Then InitCOMNow = True
        End If

        If InitCOMNow And (String.IsNullOrEmpty(SerialPort) = False) Then
            COM_port = New IO.Ports.SerialPort(SerialPort, 9600, IO.Ports.Parity.None, 8, IO.Ports.StopBits.One)
            COM_port.ReadTimeout = 1000
            COM_port.Handshake = IO.Ports.Handshake.None
            AddHandler COM_port.DataReceived, New IO.Ports.SerialDataReceivedEventHandler(AddressOf COM_port_DataReceived)
            Try
                COM_port.Open()
            Catch ex As Exception
                'COM port not available ...
            End Try
        End If

    End Sub

    '''<summary>Set voltage.</summary>
    Public Function SetU(ByVal U As Double) As Boolean
        If GetAnswer("VOLT" & Format(U * 10, "000") & StopChars).EndsWith("OK") Then Return True Else Return False
    End Function

    '''<summary>Set current.</summary>
    Public Function SetI(ByVal I As Double) As Boolean
        If GetAnswer("CURR" & Format(I * 10, "000") & StopChars).EndsWith("OK") Then Return True Else Return False
    End Function

    '''<summary>Set output state.</summary>
    Public Function SetState(ByVal State As Boolean) As Boolean
        MyOutput = CStr(IIf(State = True, "ON", "OFF")).Trim
        If GetAnswer("SOUT" & CStr(IIf(State = True, "0", "1")).Trim).EndsWith("OK") Then Return True Else Return False
    End Function

    Public Sub Update()
        'Read all answers
        Dim GMAX_Ans As String = GetAnswer(Commands.GMAX)
        Dim GETD_Ans As String = GetAnswer(Commands.GETD)
        'Parse answers
        If GMAX_Ans.EndsWith("OK") Then
            MyU_max = Val(GMAX_Ans.Substring(0, 3)) / 10
            MyI_max = Val(GMAX_Ans.Substring(3, 3)) / 10
        End If
        'Parse answers
        If GETD_Ans.EndsWith("OK") Then
            MyU = Val(GETD_Ans.Substring(0, 4)) / 100
            MyI = Val(GETD_Ans.Substring(4, 4)) / 100
            MyCVCCMode = CStr(IIf(GETD_Ans.Substring(8, 1) = "0", "CV", "CChar"))
        End If
        MyLastUpdate = Now
    End Sub

    '''<summary>Send command to dome.</summary>
    '''<param name="command">Command to send.</param>
    '''<remarks>Taken from Baader driver.</remarks>
    Private Function GetAnswer(ByVal command As String) As String
        ReadBuffer = String.Empty
        If COM_port.IsOpen Then
            COM_port.Write(command & StopChars)
            Do
                System.Threading.Thread.Sleep(10)
            Loop Until (ReadBuffer.EndsWith("OK" & StopChars))
            Return ReadBuffer.Substring(0, ReadBuffer.Length - StopChars.Length)
        Else
            Return Nothing
        End If
    End Function

    Private Sub COM_port_DataReceived(ByVal sender As Object, e As IO.Ports.SerialDataReceivedEventArgs)
        Dim str As String = CType(sender, IO.Ports.SerialPort).ReadExisting
        For i As Integer = 0 To str.Length - 1
            ReadBuffer &= str.Chars(i)
        Next i
    End Sub

End Class