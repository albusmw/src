Option Explicit On
Option Strict On
Imports System.Dynamic

'''<summary>Class to handle events from MIDI controlers.</summary>
'''<see cref="http://www.codeproject.com/Articles/814885/MIDI-monitor-written-in-Visual-Basic"/>
Public Class cMIDIMonitor

    Const DLLName As String = "winmm.dll"
    Const DLLCharSet As Runtime.InteropServices.CharSet = Runtime.InteropServices.CharSet.Auto

    Private hMidiIn As IntPtr
    Private AsyncOpHandler As ComponentModel.AsyncOperation

    '''<summary>Generic controler messages.</summary>
    Public Event Message(ByVal Message As String)
    Private Sub MessageRaiser(ByVal Message As Object)
        RaiseEvent Message(CType(Message, String))
    End Sub

    '''<summary>New data from the controls.</summary>
    Public Event Data(ByVal Channel As Integer, ByVal Value As Integer)
    Private Sub DataRaiser(ByVal Data As Object)
        RaiseEvent Data(CType(CType(Data, Object())(0), Integer), CType(CType(Data, Object())(1), Integer))
    End Sub

    '''<summary>New data from the controls.</summary>
    Public Event Increment(ByVal Channel As Integer, ByVal Delta As Integer)
    Private Sub IncrementRaiser(ByVal Data As Object)
        RaiseEvent Increment(CType(CType(Data, Object())(0), Integer), CType(CType(Data, Object())(1), Integer))
    End Sub

    '''<summary>Verbose data for logging.</summary>
    Public Event VerbLog(ByVal Text As String)
    Private Sub VerbLogRaiser(ByVal Text As Object)
        RaiseEvent VerbLog(CType(Text, String))
    End Sub

    Dim MonitorActive As Boolean = False

    Private CurrentChannelValues As New Dictionary(Of Integer, Integer)
    Private LastChannelValues As New Dictionary(Of Integer, Integer)
    Private LastOutputData As New Dictionary(Of Integer, Integer)

    '''<summary>Minimum channel value that is send by the MIDI device.</summary>
    Private MinChannelValue As Integer = 0
    '''<summary>Maximum channel value that is send by the MIDI device.</summary>
    Private MaxChannelValue As Integer = 127
    Private MidChannelValue As Integer = 64

    Public Enum MMRESULT As Integer
        NOERROR = 0
        [ERROR]
        BADDEVICEID
        NOTENABLED
        ALLOCATED
        INVALHANDLE
        NODRIVER
        NOMEM
        NOTSUPPORTED
        BADERRNUM
        INVALFLAG
        INVALPARAM
        HANDLEBUSY
        INVALIDALIAS
        BADDB
        KEYNOTFOUND
        READERROR
        WRITEERROR
        DELETEERROR
        VALNOTFOUND
        NODRIVERCB
    End Enum

    ''' <summary>The MIDIINCAPS structure describes the capabilities of a MIDI input device.</summary>
    Public Structure MIDIINCAPS
        ''' <summary>Manufacturer identifier of the device driver for the MIDI input device. Manufacturer identifiers are defined in Manufacturer and Product Identifiers.</summary>
        Public wMid As Int16
        ''' <summary>Product identifier of the MIDI input device. Product identifiers are defined in Manufacturer and Product Identifiers.</summary>
        Public wPid As Int16
        ''' <summary>Version number of the device driver for the MIDI input device. The high-order byte is the major version number, and the low-order byte is the minor version number.</summary>
        Public vDriverVersion As Integer
        ''' <summary>Product name in a null-terminated string.</summary>
        <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=32)> Dim szPname As String
        ''' <summary>Reserved; must be zero.</summary>
        Public dwSupport As Integer
    End Structure

    ''' <summary>The midiInGetNumDevs function retrieves the number of MIDI input devices in the system.</summary>
    ''' <returns>Returns the number of MIDI input devices present in the system. A return value of zero means that there are no devices (not that there is no error).</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiingetnumdevs"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInGetNumDevs() As UInt32
    End Function

    ''' <summary>The midiInGetDevCaps function determines the capabilities of a specified MIDI input device.</summary>
    ''' <param name="uDeviceID">Identifier of the MIDI input device. The device identifier varies from zero to one less than the number of devices present. This parameter can also be a properly cast device handle.</param>
    ''' <param name="lpCaps">Pointer to a MIDIINCAPS structure that is filled with information about the capabilities of the device.</param>
    ''' <param name="cbmic">Size, in bytes, of the MIDIINCAPS structure. Only cbMidiInCaps bytes (or less) of information is copied to the location pointed to by lpMidiInCaps. If cbMidiInCaps is zero, nothing is copied, and the function returns MMSYSERR_NOERROR.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiingetdevcaps"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInGetDevCaps(ByVal uDeviceID As UInt32, ByRef lpCaps As MIDIINCAPS, ByVal cbmic As UInt32) As MMRESULT
    End Function

    ''' <summary>The midiInOpen function opens a specified MIDI input device.</summary>
    ''' <param name="phmi">Pointer to an HMIDIIN handle. This location is filled with a handle identifying the opened MIDI input device. The handle is used to identify the device in calls to other MIDI input functions.</param>
    ''' <param name="uDeviceID">Identifier of the MIDI input device to be opened.</param>
    ''' <param name="dwCallback">Pointer to a callback function, a thread identifier, or a handle of a window called with information about incoming MIDI messages. For more information on the callback function, see MidiInProc.</param>
    ''' <param name="dwInstance">User instance data passed to the callback function. This parameter is not used with window callback functions or threads.</param>
    ''' <param name="fdwOpen">Callback flag for opening the device and, optionally, a status flag that helps regulate rapid data transfers.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinopen"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInOpen(ByRef phmi As IntPtr, ByVal uDeviceID As UInt32, ByVal dwCallback As MidiInCallback, ByVal dwInstance As Integer, ByVal fdwOpen As Integer) As MMRESULT
    End Function

    ''' <summary>The midiInStart function starts MIDI input on the specified MIDI input device.</summary>
    ''' <param name="hmi">Handle to the MIDI input device.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinstart"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInStart(ByVal hmi As IntPtr) As MMRESULT
    End Function

    ''' <summary>The midiInStop function stops MIDI input on the specified MIDI input device.</summary>
    ''' <param name="hmi">Handle to the MIDI input device.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinstop"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInStop(ByVal hmi As IntPtr) As MMRESULT
    End Function

    ''' <summary>The midiInReset function stops input on a given MIDI input device.</summary>
    ''' <param name="hmi">Handle to the MIDI input device.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinreset"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInReset(ByVal hMidiIn As IntPtr) As MMRESULT
    End Function

    ''' <summary>The midiInClose function closes the specified MIDI input device.</summary>
    ''' <param name="hmi">Handle to the MIDI input device.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinclose"/>
    <Runtime.InteropServices.DllImport(DLLName, SetLastError:=True, CharSet:=DLLCharSet)>
    Public Shared Function midiInClose(ByVal hMidiIn As IntPtr) As MMRESULT

    End Function

    Public Delegate Function MidiInCallback(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As UInt32, ByVal dwParam2 As Integer) As Integer
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
    Public Const CALLBACK_FUNCTION As Integer = &H30000
    Public Const MIDI_IO_STATUS = &H20




    '''<summary>Number of found MIDI devices.</summary>
    Public ReadOnly Property MIDIDeviceCount As Integer
        Get
            If IsNothing(MyMIDIDevices) = True Then Return 0 Else Return MyMIDIDevices.Length
        End Get
    End Property

    '''<summary>List of all available MIDI devices.</summary>
    Public ReadOnly Property MIDIDevices As String()
        Get
            Return MyMIDIDevices
        End Get
    End Property
    Private MyMIDIDevices As String() = {}

    '''<summary>Constructor that evaluated all connected MIDI devices.</summary>
    Public Sub New()

        'Exit on no available devices
        Dim NumDevs As UInt32 = midiInGetNumDevs()
        If NumDevs = 0 Then
            Exit Sub
        Else
            'Iterate over all devices
            ReDim MyMIDIDevices(CInt(NumDevs) - 1)
            For DevCnt As UInt32 = 0 To CUInt(NumDevs - 1)
                Dim InCaps As New MIDIINCAPS
                midiInGetDevCaps(DevCnt, InCaps, CUInt(Len(InCaps)))
                MyMIDIDevices(CInt(DevCnt)) = InCaps.szPname.Trim
            Next DevCnt
        End If

        AsyncOpHandler = ComponentModel.AsyncOperationManager.CreateOperation(Nothing)

    End Sub

    Function MidiInProc(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As UInt32, ByVal dwParam2 As Integer) As Integer

        If MonitorActive = True Then

            'Get the hex message
            Dim Message As String = Hex(dwParam1).PadLeft(6, CChar("0")) & Hex(dwParam2).PadLeft(6, CChar("0"))
            AsyncOpHandler.Post(New Threading.SendOrPostCallback(AddressOf VerbLogRaiser), CType(Message, Object))

            If MyMIDIDevices(0) = "Arturia" Then
                Decode_Arturia(Message)
            Else

                Dim DecodedData As Integer = 0
                Dim Channel As Integer = -1
                Dim MessageHeader As String = Message.Substring(2, 2)
                Dim IgnoreMessage As Boolean = False

                Select Case MessageHeader
                    Case "0A"
                        Dim Decoded As String = "Fader  =" & Val("&H" & Message.Substring(0, 2))

                    Case "0B"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 1
                    Case "0C"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 2
                    Case "0D"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 3
                    Case "0E"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 4
                    Case "0F"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 5
                    Case "10"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 6
                    Case "11"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 7
                    Case "12"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 8

                    Case "18"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 1
                        If DecodedData = MaxChannelValue Then DecodedData = -1              'Press
                        If DecodedData = MinChannelValue Then IgnoreMessage = True
                    Case "19"
                        DecodedData = CInt("&H" & Message.Substring(0, 2)) : Channel = 2
                        If DecodedData = MaxChannelValue Then DecodedData = -1              'Press
                        If DecodedData = MinChannelValue Then IgnoreMessage = True

                End Select

                If Not IgnoreMessage Then

                    'Add new channel / store
                    If CurrentChannelValues.ContainsKey(Channel) = False Then
                        CurrentChannelValues.Add(Channel, DecodedData)
                        LastChannelValues.Add(Channel, DecodedData)
                        LastOutputData.Add(Channel, 0)
                    Else
                        If DecodedData <> -1 Then
                            LastChannelValues(Channel) = CurrentChannelValues(Channel)
                            CurrentChannelValues(Channel) = DecodedData
                        End If
                    End If

                    'React on "press"
                    Dim CurrentData As Integer = -1
                    Dim Increment As Integer = 0
                    If DecodedData = -1 Then
                        CurrentData = MidChannelValue
                    Else
                        'Realize an "endless rotaty"
                        Increment = CurrentChannelValues(Channel) - LastChannelValues(Channel)

                        If Increment = 0 Then
                            If DecodedData = MinChannelValue Then Increment = -1
                            If DecodedData = MaxChannelValue Then Increment = +1
                        End If
                        CurrentData = LastOutputData(Channel) + Increment

                    End If

                    'Store last output data
                    If LastOutputData.ContainsKey(Channel) = False Then LastOutputData.Add(Channel, CurrentData) Else LastOutputData(Channel) = CurrentData
                    AsyncOpHandler.Post(New Threading.SendOrPostCallback(AddressOf IncrementRaiser), New Object() {Channel, Increment})

                    Dim Time As String = Format(dwParam2 / 1000, "000.000")

                    'RaiseEvent NewMessage(Format(hMidiIn, "000000") & ":" & Format(wMsg, "000000") & ":" & Message & ":" & " -> " & "...")

                    AsyncOpHandler.Post(New Threading.SendOrPostCallback(AddressOf DataRaiser), New Object() {Channel, CurrentData})

                End If

            End If

        End If

        Return 0

    End Function

    Public Function SelectMidiDevice(ByVal DeviceID As UInt32) As Boolean
        Dim OpenErr As MMRESULT = midiInOpen(hMidiIn, DeviceID, ptrCallback, 0, CALLBACK_FUNCTION Or MIDI_IO_STATUS)
        If OpenErr <> MMRESULT.NOERROR Then Return False
        Dim StartErr As MMRESULT = midiInStart(hMidiIn)
        If StartErr <> MMRESULT.NOERROR Then Return False
        MonitorActive = True
        Return True
    End Function

    Public Sub StartMonitor()
        midiInStart(hMidiIn)
        MonitorActive = True
    End Sub

    Public Sub StopMonitor()
        midiInStop(hMidiIn)
        MonitorActive = False
    End Sub

    Private Sub Disconnect()
        MonitorActive = False
        midiInStop(hMidiIn)
        midiInReset(hMidiIn)
    End Sub

    Private Sub Decode_Arturia(ByVal Message As String)
        Dim UpDown As Int16 = CShort("&H" & Message.Substring(0, 2))
        If UpDown >= 64 Then UpDown = CShort(UpDown - 128)
        Dim Channel As Integer = CInt("&H" & Message.Substring(2, 2))
        AsyncOpHandler.Post(New Threading.SendOrPostCallback(AddressOf DataRaiser), New Object() {Channel, UpDown})
    End Sub

End Class