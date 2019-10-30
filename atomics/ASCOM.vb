Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    Public Class ASCOM

        Public Shared Function TelescopeProperties(ByRef Telescope As Global.ASCOM.DriverAccess.Telescope, ByVal SelectedTelescope As String) As List(Of String)

            Dim RetVal As New List(Of String)
            RetVal.Add("Properties of <" & SelectedTelescope & ">:")
            Dim TelescopeType As Type = Telescope.GetType()
            Dim AllProperties As System.Reflection.PropertyInfo() = TelescopeType.GetProperties
            Dim PropertyLog As New List(Of String)
            For Each Prop As System.Reflection.PropertyInfo In AllProperties
                Try
                    'Hide some ...
                    Dim Hide As Boolean = False
                    Select Case Prop.Name.ToUpper
                        'Case "ImageArray".ToUpper, "ImageArrayVariant".ToUpper : Hide = True
                        'Case "LastExposureDuration".ToUpper, "LastExposureStartTime".ToUpper : Hide = True
                    End Select
                    If Not Hide Then
                        'Process different property types
                        Dim PropValue As New List(Of String)
                        Select Case Prop.PropertyType.Name.ToUpper
                            Case "ArrayList".ToUpper
                                For Each Entry As Object In CType(Prop.GetValue(Telescope, Nothing), ArrayList)
                                    PropValue.Add(CStr(Entry))
                                Next Entry
                            Case Else
                                PropValue.Add(DecodeProperty(Telescope, Prop))
                        End Select
                        'Display with access methods
                        If Prop.CanWrite = True And Prop.CanRead Then
                            PropertyLog.Add("  (R/W) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                        Else
                            If Prop.CanRead = True Then
                                PropertyLog.Add("  (R/-) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                            Else
                                PropertyLog.Add("  (-/W) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                            End If
                        End If
                    End If
                Catch ex As Exception
                    If IsNothing(ex.InnerException) = False Then
                        If ex.InnerException.GetType Is GetType(Global.ASCOM.PropertyNotImplementedException) Then
                            PropertyLog.Add("  (-/-) " & Prop.Name.PadRight(30) & ": --NOT IMPLEMENTED--")
                        Else
                            PropertyLog.Add("  (?/?) " & Prop.Name.PadRight(30) & ": <ASCOM Error: " & ex.InnerException.Message & ">")
                        End If
                    Else
                        PropertyLog.Add("  (?/?) " & Prop.Name.PadRight(30) & ": <" & ex.Message & ">")
                    End If
                End Try
            Next Prop
            PropertyLog.Sort()                                          'Sort alphabetically
            For Each Entry As String In PropertyLog
                RetVal.Add(Entry)
            Next Entry
            RetVal.Add("=============================================")

            Return RetVal

        End Function

        Public Shared Function CameraProperties(ByRef Camera As Global.ASCOM.DriverAccess.Camera, ByVal SelectedCamera As String) As List(Of String)

            Dim RetVal As New List(Of String)
            RetVal.Add("Properties of <" & SelectedCamera & ">:")
            Dim CamType As Type = Camera.GetType()
            Dim AllProperties As System.Reflection.PropertyInfo() = CamType.GetProperties
            Dim PropertyLog As New List(Of String)
            For Each Prop As System.Reflection.PropertyInfo In AllProperties
                Try
                    'Hide some ...
                    Dim Hide As Boolean = False
                    Select Case Prop.Name.ToUpper
                        Case "ImageArray".ToUpper, "ImageArrayVariant".ToUpper : Hide = True
                        Case "LastExposureDuration".ToUpper, "LastExposureStartTime".ToUpper : Hide = True
                    End Select
                    If Not Hide Then
                        'Process different property types
                        Dim PropValue As New List(Of String)
                        Select Case Prop.PropertyType.Name.ToUpper
                            Case "ArrayList".ToUpper
                                For Each Entry As Object In CType(Prop.GetValue(Camera, Nothing), ArrayList)
                                    PropValue.Add(CStr(Entry))
                                Next Entry
                            Case Else
                                PropValue.Add(DecodeProperty(Camera, Prop))
                        End Select
                        'Display with access methods
                        If Prop.CanWrite = True And Prop.CanRead Then
                            PropertyLog.Add("  (R/W) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                        Else
                            If Prop.CanRead = True Then
                                PropertyLog.Add("  (R/-) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                            Else
                                PropertyLog.Add("  (-/W) " & Prop.Name.PadRight(30) & ":" & Join(PropValue.ToArray, "|"))
                            End If
                        End If
                    End If
                Catch ex As Exception
                    If IsNothing(ex.InnerException) = False Then
                        If ex.InnerException.GetType Is GetType(Global.ASCOM.PropertyNotImplementedException) Then
                            PropertyLog.Add("  (-/-) " & Prop.Name.PadRight(30) & ": --NOT IMPLEMENTED--")
                        Else
                            PropertyLog.Add("  (?/?) " & Prop.Name.PadRight(30) & ": <ASCOM Error: " & ex.InnerException.Message & ">")
                        End If
                    Else
                        PropertyLog.Add("  (?/?) " & Prop.Name.PadRight(30) & ": <" & ex.Message & ">")
                    End If
                End Try
            Next Prop
            PropertyLog.Sort()                                          'Sort alphabetically
            For Each Entry As String In PropertyLog
                RetVal.Add(Entry)
            Next Entry
            RetVal.Add("=============================================")

            Return RetVal

        End Function

        Public Shared Function DecodeProperty(ByRef Camera As Global.ASCOM.DriverAccess.Camera, ByRef Prop As System.Reflection.PropertyInfo) As String
            Select Case Prop.PropertyType.Name.ToUpper
                Case "CameraStates".ToUpper
                    Return "<" & [Enum].GetName(GetType(Global.ASCOM.DeviceInterface.CameraStates), Prop.GetValue(Camera, Nothing)) & ">"
                Case "SensorType".ToUpper
                    Return "<" & [Enum].GetName(GetType(Global.ASCOM.DeviceInterface.SensorType), Prop.GetValue(Camera, Nothing)) & ">"
                Case Else
                    Return CStr(Prop.GetValue(Camera, Nothing)) '& " (" & Prop.PropertyType.Name & ")"
            End Select
        End Function

        Public Shared Function DecodeProperty(ByRef Telescope As Global.ASCOM.DriverAccess.Telescope, ByRef Prop As System.Reflection.PropertyInfo) As String
            Select Case Prop.PropertyType.Name.ToUpper
                Case "CameraStates".ToUpper
                    Return "<" & [Enum].GetName(GetType(Global.ASCOM.DeviceInterface.CameraStates), Prop.GetValue(Telescope, Nothing)) & ">"
                Case "SensorType".ToUpper
                    Return "<" & [Enum].GetName(GetType(Global.ASCOM.DeviceInterface.SensorType), Prop.GetValue(Telescope, Nothing)) & ">"
                Case Else
                    Return CStr(Prop.GetValue(Telescope, Nothing)) '& " (" & Prop.PropertyType.Name & ")"
            End Select
        End Function

    End Class

End Namespace