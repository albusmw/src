Option Explicit On
Option Strict On

'''<summary>Specific ZWO ASI functions.</summary>
Public Class cZWOASI

    Public Event Info(ByVal Text As String)
    Public Event CoolingHistory(ByVal Time As Double(), ByVal Temperature As Double(), ByVal CoolingPower As Double())

    '''<summary>Supported bin modes.</summary>
    Public Function SupportedBins(ByVal Values As Integer()) As String
        Dim RetVal As New List(Of String)
        For Each Entry As Integer In Values
            If Entry <> 0 Then RetVal.Add(Entry.ValRegIndep & "x" & Entry.ValRegIndep)
        Next Entry
        Return Join(RetVal.ToArray, "|")
    End Function

    '''<summary>Supported video modes.</summary>
    Public Function SupportedVideoFormat(ByVal Values As ZWO.ASICameraDll.ASI_IMG_TYPE()) As String
        Dim RetVal As New List(Of String)
        For Each Entry As ZWO.ASICameraDll.ASI_IMG_TYPE In Values
            RetVal.Add(Entry.ToString.Trim)
        Next Entry
        Return Join(RetVal.ToArray, "|")
    End Function

    '''<summary>Cool the camera to a certain temperature.</summary>
    '''<param name="Connection">Camera that shall be cooled.</param>
    '''<param name="TargetTemperature">Temperature the camera should be at.</param>
    '''<returns>TRUE if camera is colled at set point.</returns>
    Public Function CoolASICamera(ByRef CamHandle As Integer, ByVal TargetTemperature As Double) As Boolean
        Return CoolASICamera(CamHandle, TargetTemperature, 0.2, 60)
    End Function

    '''<summary>Cool the camera to a certain temperature.</summary>
    '''<param name="Connection">Camera that shall be cooled.</param>
    '''<param name="TargetTemperature">Temperature the camera should be at.</param>
    '''<returns>TRUE if camera is colled at set point.</returns>
    Public Function CoolASICamera(ByRef CamHandle As Integer, ByVal TargetTemperature As Double, ByVal MaxTolerance As Double, ByVal SecondsToBeIn As Integer) As Boolean

        Dim PollSpeed As Integer = 200              'Pollspeed [ms]

        Dim FirstTimeInRange As DateTime = Nothing
        Dim InRangeFor As Double = 0

        ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_COOLER_ON, 1)
        ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_TARGET_TEMP, CInt(Math.Round(TargetTemperature)))

        Dim PlotTime As New List(Of Double)
        Dim TPlot As New List(Of Double)
        Dim PowerPlot As New List(Of Double)

        Dim CurrentTime As Double = 0
        Do
            System.Threading.Thread.Sleep(PollSpeed)
            Dim CurrentTemperature As Double = ZWO.ASICameraDll.ASIGetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_TEMPERATURE) / 10
            If System.Math.Abs(TargetTemperature - CurrentTemperature) <= MaxTolerance Then
                If FirstTimeInRange = Nothing Then FirstTimeInRange = Now
                InRangeFor = Now.Subtract(FirstTimeInRange).TotalSeconds
                If InRangeFor >= SecondsToBeIn Then Return True
            Else
                FirstTimeInRange = Nothing
                InRangeFor = 0
            End If
            PlotTime.Add(CurrentTime)
            TPlot.Add(CurrentTemperature)
            PowerPlot.Add(ZWO.ASICameraDll.ASIGetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_COOLER_POWER_PERC))
            RaiseEvent Info("Target: " & TargetTemperature.ValRegIndep & " °C / Current: " & CurrentTemperature.ValRegIndep & ": In range for " & InRangeFor.ValRegIndep & " seconds")
            RaiseEvent CoolingHistory(PlotTime.ToArray, TPlot.ToArray, PowerPlot.ToArray)
            CurrentTime += (PollSpeed / 1000)
        Loop

    End Function

End Class