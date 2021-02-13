Option Explicit On
Option Strict On

Module ZWO_ASI_FromQHYCapture

    'Code is taken from QHYCapture

    Private Sub tsmiASIZWO_Click(sender As Object, e As EventArgs) Handles tsmiASIZWO.Click

        'Get number of cameras, if there is at least one continue
        Dim Cameras As Integer = ZWO.ASICameraDll.ASIGetNumOfConnectedCameras()

        Const DoNotSet As Integer = Integer.MinValue

        If Cameras > 0 Then

            Dim CamHandle As Integer = 0
            Log("Opening first camera ...")

            Dim CameraInfo As ZWO.ASICameraDll.ASI_CAMERA_INFO = Nothing
            Dim CameraID As ZWO.ASICameraDll.ASI_ID = Nothing ': CameraID.ID = New Byte() {Asc("M"), Asc("A"), Asc("R"), Asc("T"), Asc("I"), Asc("N"), Asc("_"), Asc("W")}
            Dim NumberOfControls As Integer = 0

            CallOK(ZWO.ASICameraDll.ASIOpenCamera(CamHandle))
            CallOK(ZWO.ASICameraDll.ASIGetCameraProperty(CameraInfo, CamHandle))

            'Get (and set before) the camera ID - this should only be done once for each camera ...
            'CallOK(ZWO.ASICameraDll.ASISetID(CamHandle, CameraID))
            CallOK(ZWO.ASICameraDll.ASIGetID(CamHandle, CameraID))

            'Display all camera info elements
            Log("Camera Info for <" & CameraID.IDAsString & ">:")
            Log("  " & "Name".PadRight(20) & ": " & CameraInfo.NameAsString)
            Log("  " & "CameraID".PadRight(20) & ": " & CameraInfo.CameraID)
            Log("  " & "MaxHeight".PadRight(20) & ": " & CameraInfo.MaxHeight)
            Log("  " & "MaxWidth".PadRight(20) & ": " & CameraInfo.MaxWidth)
            Log("  " & "IsColorCam".PadRight(20) & ": " & CameraInfo.IsColorCam)
            Log("  " & "BayerPattern".PadRight(20) & ": " & CameraInfo.BayerPattern)
            Log("  " & "SupportedBins".PadRight(20) & ": " & ZWOASI.SupportedBins(CameraInfo.SupportedBins))
            Log("  " & "SupportedVideoFormat".PadRight(20) & ": " & ZWOASI.SupportedVideoFormat(CameraInfo.SupportedVideoFormat))
            Log("  " & "PixelSize".PadRight(20) & ": " & CameraInfo.PixelSize)
            Log("  " & "MechanicalShutter".PadRight(20) & ": " & CameraInfo.MechanicalShutter)
            Log("  " & "ST4Port".PadRight(20) & ": " & CameraInfo.ST4Port)
            Log("  " & "IsCoolerCam".PadRight(20) & ": " & CameraInfo.IsCoolerCam)
            Log("  " & "IsUSB3Host".PadRight(20) & ": " & CameraInfo.IsUSB3Host)
            Log("  " & "IsUSB3Camera".PadRight(20) & ": " & CameraInfo.IsUSB3Camera)
            Log("  " & "ElecPerADU".PadRight(20) & ": " & CameraInfo.ElecPerADU)
            Log("  " & "BitDepth".PadRight(20) & ": " & CameraInfo.BitDepth)
            Log("  " & "IsTriggerCam".PadRight(20) & ": " & CameraInfo.IsTriggerCam)
            Log("=======================================")

            'Open camera
            CallOK(ZWO.ASICameraDll.ASIInitCamera(CamHandle))

            'Get different number of controls
            ZWO.ASICameraDll.ASIGetNumOfControls(0, NumberOfControls)
            Log(NumberOfControls.ValRegIndep & " controls")

            'Read out every control that are available and set to default
            Log("ControlCaps:")
            For ControlIdx As Integer = 0 To NumberOfControls - 1
                Dim ControlCap As ZWO.ASICameraDll.ASI_CONTROL_CAPS = Nothing
                ZWO.ASICameraDll.ASIGetControlCaps(CamHandle, ControlIdx, ControlCap)
                Log("  " & ControlCap.NameAsString.PadRight(30) & ": " & TabFormat(ControlCap.MinValue) & " ... " & TabFormat(ControlCap.MaxValue) & ",default: " & TabFormat(ControlCap.DefaultValue) & "(" & ControlCap.DescriptionAsString & ")")
                ZWO.ASICameraDll.ASISetControlValue(CamHandle, ControlCap.ControlType, ControlCap.DefaultValue)
            Next ControlIdx
            Log("=======================================")

            'Read out all set values
            Log("ControlValues:")
            For Each X As ZWO.ASICameraDll.ASI_CONTROL_TYPE In [Enum].GetValues(GetType(ZWO.ASICameraDll.ASI_CONTROL_TYPE))
                Log("  " & X.ToString.Trim.PadRight(30) & ": " & ZWO.ASICameraDll.ASIGetControlValue(CamHandle, X))
            Next X
            Log("=======================================")

            'Read out all set values
            Log("SpecialValues:")
            Dim Offset_HighestDR As Integer : Dim Offset_UnityGain As Integer : Dim Gain_LowestRN As Integer : Dim Offset_LowestRN As Integer
            ZWO.ASICameraDll.ASIGetGainOffset(CamHandle, Offset_HighestDR, Offset_UnityGain, Gain_LowestRN, Offset_LowestRN)
            Log("  " & "Offset configuration:")
            Log("  " & "  HighestDR         : " & Offset_HighestDR.ValRegIndep)
            Log("  " & "  UnityGain         : " & Offset_UnityGain.ValRegIndep)
            Log("  " & "  LowestRN          : " & Offset_LowestRN.ValRegIndep)
            Log("  " & "    @               : " & Gain_LowestRN.ValRegIndep)
            Log("=======================================")

            'Run exposure sequence configured
            Log("Exposing ...")
            Dim ROIWidth As Integer = -1
            Dim ROIHeight As Integer = -1
            Dim ROIBin As Integer = -1
            Dim ROIImgType As ZWO.ASICameraDll.ASI_IMG_TYPE = ZWO.ASICameraDll.ASI_IMG_TYPE.ASI_IMG_END
            Dim StartPosX As Integer = -1
            Dim StartPosY As Integer = -1

            CallOK(ZWO.ASICameraDll.ASISetROIFormat(CamHandle, CameraInfo.MaxWidth, CameraInfo.MaxHeight, 1, ZWO.ASICameraDll.ASI_IMG_TYPE.ASI_IMG_RAW16))
            CallOK(ZWO.ASICameraDll.ASIGetROIFormat(CamHandle, ROIWidth, ROIHeight, ROIBin, ROIImgType))
            CallOK(ZWO.ASICameraDll.ASISetStartPos(CamHandle, 0, 0))
            CallOK(ZWO.ASICameraDll.ASIGetStartPos(CamHandle, StartPosX, StartPosY))

            'Prepare logging
            Dim Path As String = System.IO.Path.Combine(M.DB.StoragePath, M.Meta.GUID)
            If System.IO.Directory.Exists(Path) = False Then System.IO.Directory.CreateDirectory(Path)
            Dim CSVLogPath As String = System.IO.Path.Combine(Path, "ZWO_Statisic_Log.csv")
            System.IO.File.WriteAllText(System.IO.Path.Combine(Path, "InitialLog.log"), tbLogOutput.Text)

            'Prepare buffers
            Dim CamRawBuffer((CameraInfo.MaxWidth * CameraInfo.MaxHeight * 2) - 1) As Byte
            Dim CamRawGAC As Runtime.InteropServices.GCHandle = Runtime.InteropServices.GCHandle.Alloc(CamRawBuffer, Runtime.InteropServices.GCHandleType.Pinned)
            Dim CamRawPtr As IntPtr = System.Runtime.InteropServices.Marshal.UnsafeAddrOfPinnedArrayElement(CamRawBuffer, 0)
            Dim CamRawBufferBytes As Integer = CInt(CamRawBuffer.LongLength * 2)

            'Prepare statistics
            Dim SingleStatCalc As New AstroNET.Statistics(M.DB.IPP)
            Dim ExpCounter As Integer = 0
            Dim SweepCount As Integer = 10

            Dim CSVLog As New Ato.cCSVBuilder

            For Each TargetTemp As Integer In New Integer() {DoNotSet}

                'Cooling
                If TargetTemp <> Integer.MinValue Then ZWOASI.CoolASICamera(CamHandle, TargetTemp, 0.2, 120)

                For Each ExpTimeToSet As Integer In New Integer() {1}

                    CallOK(ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_EXPOSURE, ExpTimeToSet))

                    For Each GainToSet As Integer In New Integer() {1}

                        If GainToSet <> DoNotSet Then CallOK(ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_GAIN, GainToSet))

                        For Each GammaToSet As Integer In New Integer() {50}

                            If GammaToSet <> DoNotSet Then CallOK(ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_GAMMA, GammaToSet))

                            For Each BrightnessToSet As Integer In New Integer() {50}

                                If BrightnessToSet <> DoNotSet Then CallOK(ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_BRIGHTNESS, BrightnessToSet))

                                'Configure all control values
                                ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_WB_B, 0)
                                ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_WB_R, 0)
                                ZWO.ASICameraDll.ASISetControlValue(CamHandle, ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_MONO_BIN, 1)

                                'Compose logging info
                                Dim LogInfo As String = "T=" & (TargetTemp / 10).ValRegIndep & "|Exp=" & (ExpTimeToSet / 1000) & "ms|Gain=" & GainToSet.ValRegIndep & "|Offset=" & BrightnessToSet.ValRegIndep & "|Gamma=" & GammaToSet.ValRegIndep

                                Dim LoopStat As New AstroNET.Statistics.sStatistics

                                For LoopCnt As Integer = 1 To SweepCount

                                    'Log current parameters
                                    ExpCounter += 1

                                    CSVLog.StartRow()
                                    CSVLog.AddColumnValue("ExpCounter", ExpCounter)
                                    For Each X As ZWO.ASICameraDll.ASI_CONTROL_TYPE In [Enum].GetValues(GetType(ZWO.ASICameraDll.ASI_CONTROL_TYPE))
                                        Dim ColumnName As String = X.ToString.Replace("ASI_", String.Empty)
                                        Select Case X
                                            Case ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_EXPOSURE
                                                CSVLog.AddColumnValue(ColumnName, ZWO.ASICameraDll.ASIGetControlValue(CamHandle, X) / 1000000)      '[s]
                                            Case ZWO.ASICameraDll.ASI_CONTROL_TYPE.ASI_TEMPERATURE
                                                CSVLog.AddColumnValue(ColumnName, ZWO.ASICameraDll.ASIGetControlValue(CamHandle, X) / 10)          '[°C]
                                            Case Else
                                                CSVLog.AddColumnValue(ColumnName, ZWO.ASICameraDll.ASIGetControlValue(CamHandle, X))
                                        End Select
                                    Next X

                                    'Start test exposure
                                    Log("Exposure " & LoopCnt.ValRegIndep & "/" & SweepCount & ":" & LogInfo)
                                    Dim Ticker As New Stopwatch : Ticker.Reset() : Ticker.Start()
                                    CallOK(ZWO.ASICameraDll.ASIStartExposure(CamHandle, ZWO.ASICameraDll.ASI_BOOL.ASI_FALSE))
                                    Dim ExpStatus As ZWO.ASICameraDll.ASI_EXPOSURE_STATUS = ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_FAILED
                                    Dim ExpFailedCount As Integer = 0
                                    Do
                                        System.Threading.Thread.Sleep(1)
                                        'Application.DoEvents()

                                        If CallOK(ZWO.ASICameraDll.ASIGetExpStatus(CamHandle, ExpStatus)) = False Then
                                            Exit Do
                                        Else
                                            Select Case ExpStatus
                                                Case ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_FAILED
                                                    'Restart exposure
                                                    ExpFailedCount += 1
                                                    ZWO.ASICameraDll.ASIStopExposure(CamHandle)
                                                    ZWO.ASICameraDll.ASIStartExposure(CamHandle, ZWO.ASICameraDll.ASI_BOOL.ASI_FALSE)
                                                    Log("###### EXPOSING FAILED! ######")
                                                Case ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_IDLE
                                                    Log("Exposure idle, existing ...")
                                                Case ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_WORKING
                                                    'Still working ...
                                                Case ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_SUCCESS
                                                    Exit Do
                                            End Select
                                        End If
                                    Loop Until 1 = 0
                                    Ticker.Stop()
                                    LogTiming("Exposure duration", Ticker)

                                    'Get data - due to some unknown reason the buffer must be "X times bigger" compared to the expected size; X = 3 works ...
                                    If ExpStatus = ZWO.ASICameraDll.ASI_EXPOSURE_STATUS.ASI_EXP_SUCCESS Then

                                        'Read data
                                        Log("Reading data ...")
                                        Ticker.Reset() : Ticker.Start()
                                        CallOK(ZWO.ASICameraDll.ASIGetDataAfterExp(CamHandle, CamRawPtr, CamRawBufferBytes))
                                        Ticker.Stop()
                                        LogTiming("", Ticker)

                                        'Correct aspect and calculate statistics
                                        Log("Statistics ...")
                                        Ticker.Reset() : Ticker.Start()
                                        SingleStatCalc.DataProcessor_UInt16.ImageData(0).Data = ChangeAspectIPP(M.DB.IPP, CamRawBuffer, CameraInfo.MaxWidth, CameraInfo.MaxHeight)
                                        Dim SingleStat As New AstroNET.Statistics.sStatistics
                                        SingleStat = SingleStatCalc.ImageStatistics(SingleStat.DataMode)
                                        LoopStat = AstroNET.Statistics.CombineStatistics(SingleStat.DataMode, SingleStat, LoopStat)
                                        Ticker.Stop()
                                        LogTiming("", Ticker)

                                        'Plot histogram
                                        With M.Report
                                            .Plotter.Clear()
                                            .Plotter.PlotXvsY(.Prop.BayerPattern(0), LoopStat.BayerHistograms_Int(0, 0), New cZEDGraphService.sGraphStyle(Color.Red, 1))
                                            .Plotter.PlotXvsY(.Prop.BayerPattern(1), LoopStat.BayerHistograms_Int(0, 1), New cZEDGraphService.sGraphStyle(Color.LightGreen, 1))
                                            .Plotter.PlotXvsY(.Prop.BayerPattern(2), LoopStat.BayerHistograms_Int(1, 0), New cZEDGraphService.sGraphStyle(Color.DarkGreen, 1))
                                            .Plotter.PlotXvsY(.Prop.BayerPattern(3), LoopStat.BayerHistograms_Int(1, 1), New cZEDGraphService.sGraphStyle(Color.Blue, 1))
                                            .Plotter.PlotXvsY("Mono histo", LoopStat.MonochromHistogram_Int, New cZEDGraphService.sGraphStyle(Color.Black, 1))
                                            .Plotter.ManuallyScaleXAxis(LoopStat.MonoStatistics_Int.Min.Key, LoopStat.MonoStatistics_Int.Max.Key)

                                            .Plotter.AutoScaleYAxisLog()
                                            .Plotter.GridOnOff(True, True)
                                            .Plotter.ForceUpdate()
                                        End With


                                        'Write image data
                                        If M.DB.StoreImage = True Then
                                            Log("Storing ...")
                                            Ticker.Reset() : Ticker.Start()
                                            Dim FITSName As String = System.IO.Path.Combine(Path, "EXP_" & Format(ExpCounter, "000000000").Trim & ".fits")
                                            cFITSWriter.Write(FITSName, SingleStatCalc.DataProcessor_UInt16.ImageData(0).Data, cFITSWriter.eBitPix.Int16)
                                            Ticker.Stop()
                                            LogTiming("", Ticker)
                                            'Process.Start(FITSName)
                                        End If

                                        'Log and display statistics
                                        CSVLog.AddColumnValue("ExpFailedCount", ExpFailedCount.ValRegIndep)
                                        CSVLog.AddColumnValue("Mean", LoopStat.MonoStatistics_Int.Mean)
                                        CSVLog.AddColumnValue("Median", LoopStat.MonoStatistics_Int.Median)
                                        CSVLog.AddColumnValue("Min", LoopStat.MonoStatistics_Int.Min.Key)
                                        CSVLog.AddColumnValue("Max", LoopStat.MonoStatistics_Int.Max.Key)
                                        CSVLog.AddColumnValue("StdDev", LoopStat.MonoStatistics_Int.StdDev)
                                        CSVLog.AddColumnValue("DifferentADUValues", LoopStat.MonoStatistics_Int.DifferentADUValues)
                                        For Each Perc As Byte In LoopStat.MonoStatistics_Int.Percentile.Keys
                                            CSVLog.AddColumnValue("Perc_" & Perc.ValRegIndep, LoopStat.MonoStatistics_Int.Percentile(Perc))
                                        Next Perc

                                    End If

                                    CSVLog.AddColumnValue("Timing", Ticker.ElapsedMilliseconds)
                                    CSVLog.AddColumnValue("ExpStatus", ExpStatus.ToString.Trim)

                                    System.IO.File.WriteAllText(CSVLogPath, CSVLog.CreateCSV)

                                Next LoopCnt

                            Next BrightnessToSet

                        Next GammaToSet

                    Next GainToSet

                Next ExpTimeToSet

            Next TargetTemp

            'Release buffers
            CamRawGAC.Free()

            'Close camera
            Log("Closing camera ...")
            CallOK(ZWO.ASICameraDll.ASICloseCamera(CamHandle))
            GC.Collect()
            Log("=======================================")

        End If

    End Sub

End Module