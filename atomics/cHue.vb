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

    '''<summary>Key for this API.</summary>
    Public Property MyPersonalAppKey As String = "lQcLDFzNwxOvkEpppP6loXXRUacXRL4tIG3gGQRN"
    Public Property BridgeIP As String = "192.168.10.183"

    Private Const APIAddressTemplate As String = "http://{0}/api"
    Private Const BodyTemplate As String = "{{\""devicetype\"":\""{0}\""}}"

    Public Async Sub Go(ByVal State As Boolean)

        Dim Client As New Q42.HueApi.LocalHueClient(BridgeIP)
        Client.Initialize(MyPersonalAppKey)
        Dim AllLights As IEnumerable(Of Q42.HueApi.Light) = Await Client.GetLightsAsync

        Dim Command As New Q42.HueApi.LightCommand

        Command.On = State
        Command.Saturation = 255
        Command.Brightness = 100
        Command.ColorCoordinates = New Double() {1.0, 0.0}
        Dim Light_0 As Q42.HueApi.Models.Groups.HueResults = Await Client.SendCommandAsync(Command, New String() {AllLights(0).Id})

        Command.ColorCoordinates = New Double() {1.0, 0.0}
        Dim Light_1 As Q42.HueApi.Models.Groups.HueResults = Await Client.SendCommandAsync(Command, New String() {AllLights(1).Id})

    End Sub


End Class