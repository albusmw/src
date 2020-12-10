Option Explicit On
Option Strict On
Imports Newtonsoft.Json

Public Class cHue

    'Alternative project with "all self done":
    ''https://steemit.com/programming/@geggi632/1-how-to-connect-to-hue-bridge-or-philips-hue-c-programming

    'This project uses Q42.HueApi
    'https://github.com/Q42/Q42.HueApi
    'Install-Package Q42.HueApi -Version 3.16.0
    'Demo: https://github.com/andriks2/Andriks.HueApiDemo
    'For color conversion see https://www.codeproject.com/Articles/613798/Colorspaces-and-Conversions

    '''<summary>IP of the hue bridge.</summary>
    Public Property BridgeIP As String = "192.168.10.183"
    '''<summary>Key to access the bridge from this app.</summary>
    Public Property AppKey As String = "lQcLDFzNwxOvkEpppP6loXXRUacXRL4tIG3gGQRN"

    Private Const APIAddressTemplate As String = "http://{0}/api"
    Private Const BodyTemplate As String = "{{\""devicetype\"":\""{0}\""}}"

    ''' <summary>Set the state of all lights OFF.</summary>
    ''' <param name="State">State (FALSE=Off, TRUE=On).</param>
    ''' <param name="Saturation">Saturation.</param>
    ''' <param name="Brightness">Brightness.</param>
    Public Sub AllOff()
        AllLights(False, 0, 0, Color.Black)
    End Sub

    ''' <summary>Set the state of all lights.</summary>
    ''' <param name="State">State (FALSE=Off, TRUE=On).</param>
    ''' <param name="Saturation">Saturation.</param>
    ''' <param name="Brightness">Brightness (0...255).</param>
    Public Async Sub AllLights(ByVal State As Boolean, ByVal Saturation As Byte, ByVal Brightness As Byte, ByVal LEDColor As Color)

        Dim Client As New Q42.HueApi.LocalHueClient(BridgeIP)
        Client.Initialize(AppKey)
        Dim AllLights As IEnumerable(Of Q42.HueApi.Light) = Await Client.GetLightsAsync

        Dim Command As New Q42.HueApi.LightCommand

        Command.On = State
        Command.Saturation = Saturation
        Command.Brightness = Brightness
        Command = Q42.HueApi.ColorConverters.Original.LightCommandExtensions.SetColor(Command, New Q42.HueApi.ColorConverters.RGBColor(LEDColor.R, LEDColor.G, LEDColor.B), "LCT001")
        Dim AllLightsResult As Q42.HueApi.Models.Groups.HueResults = Await Client.SendCommandAsync(Command)

    End Sub


End Class