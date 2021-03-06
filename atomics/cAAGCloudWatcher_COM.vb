Option Explicit On
Option Strict On

'==================================================================================
' NOT FINISHED !!!!!
'==================================================================================

'''<summary>Class for direct AAG Cloud Watcher access.</summary>
'''<see cref="https://lunaticoastro.com/aag-cloud-watcher/moreinfo/"/>
Public Class cAAGCloudWatcher_COM

    Dim COM_port As IO.Ports.SerialPort
    Dim ReadBuffer As String = String.Empty
    Dim StopChars As String = Chr(13) & Chr(10)

    '''<summary>Supported commands.</summary>
    Public Class Commands
        '''<summary>All Segments will be opened.</summary>
        Public Shared ReadOnly Property [InternalName] As String = "A!"
    End Class

    Public Sub Init(ByVal SerialPort As String)

        Dim InitCOMNow As Boolean = False
        If IsNothing(COM_port) = True Then
            InitCOMNow = True
        Else
            If COM_port.IsOpen = False Then InitCOMNow = True
        End If

        If InitCOMNow Then
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

    '''<summary>Send command to dome.</summary>
    '''<param name="command">Command to send.</param>
    Public Function GetAnswer(ByVal command As String) As String
        ReadBuffer = String.Empty
        If COM_port.IsOpen Then
            COM_port.Write(command)
            Do
                System.Threading.Thread.Sleep(10)
            Loop Until ReadBuffer.Length >= 15
            Return ReadBuffer.Substring(0, ReadBuffer.Length - StopChars.Length)
        Else
            Return Nothing
        End If
    End Function

    Public Sub Parse_rx(ByVal Answer As String, ByRef Magnitude As Double, ByRef Temperature As Double)
        If Answer.Length < 55 Then Exit Sub
        Try
            Magnitude = Val(Answer.Substring(3, 5))
        Catch ex As Exception
            Magnitude = Double.NaN
        End Try
        Try
            Temperature = Val(Answer.Substring(49, 5))
        Catch ex As Exception
            Temperature = Double.NaN
        End Try
    End Sub

    Private Sub COM_port_DataReceived(ByVal sender As Object, e As IO.Ports.SerialDataReceivedEventArgs)
        Dim str As String = CType(sender, IO.Ports.SerialPort).ReadExisting
        For i As Integer = 0 To str.Length - 1
            ReadBuffer &= str.Chars(i)
        Next i
    End Sub

End Class