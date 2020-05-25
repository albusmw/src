Option Explicit On
Option Strict On

'''<summary>Class to handle events from MIDI controlers.</summary>
'''<see cref="http://www.codeproject.com/Articles/814885/MIDI-monitor-written-in-Visual-Basic"/>
Public Class cMIDIMonitor

    Public Enum MMSYSERR As Integer
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
    Public Declare Function midiInGetNumDevs Lib "winmm.dll" () As Integer

    ''' <summary>The midiInGetDevCaps function determines the capabilities of a specified MIDI input device.</summary>
    ''' <param name="uDeviceID">Identifier of the MIDI input device. The device identifier varies from zero to one less than the number of devices present. This parameter can also be a properly cast device handle.</param>
    ''' <param name="lpCaps">Pointer to a MIDIINCAPS structure that is filled with information about the capabilities of the device.</param>
    ''' <param name="cbmic">Size, in bytes, of the MIDIINCAPS structure. Only cbMidiInCaps bytes (or less) of information is copied to the location pointed to by lpMidiInCaps. If cbMidiInCaps is zero, nothing is copied, and the function returns MMSYSERR_NOERROR.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiingetdevcaps"/>
    Public Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Integer, ByRef lpCaps As MIDIINCAPS, ByVal cbmic As Integer) As MMSYSERR

    ''' <summary>The midiInOpen function opens a specified MIDI input device.</summary>
    ''' <param name="phmi">Pointer to an HMIDIIN handle. This location is filled with a handle identifying the opened MIDI input device. The handle is used to identify the device in calls to other MIDI input functions.</param>
    ''' <param name="uDeviceID">Identifier of the MIDI input device to be opened.</param>
    ''' <param name="dwCallback">Pointer to a callback function, a thread identifier, or a handle of a window called with information about incoming MIDI messages. For more information on the callback function, see MidiInProc.</param>
    ''' <param name="dwInstance">User instance data passed to the callback function. This parameter is not used with window callback functions or threads.</param>
    ''' <param name="fdwOpen">Callback flag for opening the device and, optionally, a status flag that helps regulate rapid data transfers.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinopen"/>
    Public Declare Function midiInOpen Lib "winmm.dll" (ByRef phmi As Integer, ByVal uDeviceID As Integer, ByVal dwCallback As MidiInCallback, ByVal dwInstance As Integer, ByVal fdwOpen As Integer) As MMSYSERR

    ''' <summary>The midiInStart function starts MIDI input on the specified MIDI input device.</summary>
    ''' <param name="hmi">Handle to the MIDI input device.</param>
    ''' <returns>Returns MMSYSERR_NOERROR if successful or an error otherwise.</returns>
    ''' <see cref="https://docs.microsoft.com/en-us/windows/win32/api/mmeapi/nf-mmeapi-midiinstart"/>
    Public Declare Function midiInStart Lib "winmm.dll" (ByVal hmi As Integer) As MMSYSERR
    Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer
    Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Integer) As Integer

    Public Delegate Function MidiInCallback(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As UInt32, ByVal dwParam2 As Integer) As Integer
    Public ptrCallback As New MidiInCallback(AddressOf MidiInProc)
    Public Const CALLBACK_FUNCTION As Integer = &H30000
    Public Const MIDI_IO_STATUS = &H20

    Public Delegate Sub DisplayDataDelegate(ByVal dwParam1 As UInt32)

    Private hMidiIn As Integer

    Private CurrentChannelValues As New Dictionary(Of Integer, Integer)
    Private LastChannelValues As New Dictionary(Of Integer, Integer)
    Private LastOutputData As New Dictionary(Of Integer, Integer)

    Private MinChannelValue As Integer = 0
    Private MaxChannelValue As Integer = 127
    Private MidChannelValue As Integer = 64

    '''<summary>Generic controler messages.</summary>
    Public Event NewMessage(ByVal Message As String)
    '''<summary>New data from the controls.</summary>
    Public Event NewData(ByVal Channel As Integer, ByVal Value As Integer)
    '''<summary>Verbose data for logging.</summary>
    Public Event VerboseLog(ByVal Text As String)

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
        If midiInGetNumDevs() = 0 Then Exit Sub

        'Iterate over all devices
        ReDim MyMIDIDevices(midiInGetNumDevs - 1)
        For DevCnt As Integer = 0 To midiInGetNumDevs - 1
            Dim InCaps As New MIDIINCAPS
            midiInGetDevCaps(DevCnt, InCaps, Len(InCaps))
            MyMIDIDevices(DevCnt) = InCaps.szPname.Trim
        Next DevCnt

    End Sub

    Dim StatusByte As Byte
    Dim DataByte1 As Byte
    Dim DataByte2 As Byte
    Dim MonitorActive As Boolean = False

    Public Property HideMidiSysMessages() As Boolean
        Get
            Return MyHideMidiSysMessages
        End Get
        Set(value As Boolean)
            MyHideMidiSysMessages = value
        End Set
    End Property
    Dim MyHideMidiSysMessages As Boolean = False

    Function MidiInProc(ByVal hMidiIn As Integer, ByVal wMsg As UInteger, ByVal dwInstance As Integer, ByVal dwParam1 As UInt32, ByVal dwParam2 As Integer) As Integer

        If MonitorActive = True Then

            'Get the hex message
            Dim Message As String = Hex(dwParam1).PadLeft(6, CChar("0")) & Hex(dwParam2).PadLeft(6, CChar("0"))
            RaiseEvent VerboseLog(Message)

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
                    If DecodedData = -1 Then
                        CurrentData = MidChannelValue
                    Else
                        'Realize an "endless rotaty"
                        Dim Increment As Integer = CurrentChannelValues(Channel) - LastChannelValues(Channel)

                        If Increment <> 0 Then
                            CurrentData = LastOutputData(Channel) + Increment
                        Else
                            If DecodedData = MinChannelValue Then CurrentData = LastOutputData(Channel) - 1
                            If DecodedData = MaxChannelValue Then CurrentData = LastOutputData(Channel) + 1
                        End If

                    End If

                    'Store last output data
                    If LastOutputData.ContainsKey(Channel) = False Then LastOutputData.Add(Channel, CurrentData) Else LastOutputData(Channel) = CurrentData

                    Dim Time As String = Format(dwParam2 / 1000, "000.000")

                    RaiseEvent NewMessage(Format(hMidiIn, "000000") & ":" & Format(wMsg, "000000") & ":" & Message & ":" & " -> " & "...")
                    RaiseEvent NewData(Channel, CurrentData)

                End If

            End If

        End If

        Return 0

    End Function

    Private Sub DisplayData(ByVal dwParam1 As UInt32)
        If ((HideMidiSysMessages = True) And ((dwParam1 And &HF0) = &HF0)) Then
            Exit Sub
        Else
            StatusByte = CByte((dwParam1 And &HFF))
            DataByte1 = CByte((dwParam1 And &HFF00) >> 8)
            DataByte2 = CByte((dwParam1 And &HFF0000) >> 16)
            RaiseEvent NewMessage(String.Format("{0:X2} {1:X2} {2:X2}{3}", StatusByte, DataByte1, DataByte2, vbCrLf))
        End If
    End Sub

    Public Function SelectMidiDevice(ByVal DeviceID As Integer) As Boolean
        Dim OpenErr As MMSYSERR = midiInOpen(hMidiIn, DeviceID, ptrCallback, 0, CALLBACK_FUNCTION Or MIDI_IO_STATUS)
        If OpenErr <> MMSYSERR.NOERROR Then Return False
        Dim StartErr As MMSYSERR = midiInStart(hMidiIn)
        If StartErr <> MMSYSERR.NOERROR Then Return False
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

    Private Sub Decode_XTouch()

    End Sub

    Private Sub Decode_Arturia(ByVal Message As String)
        Dim UpDown As Int16 = CShort("&H" & Message.Substring(0, 2))
        If UpDown >= 64 Then UpDown = CShort(UpDown - 128)
        Dim Channel As Integer = CInt("&H" & Message.Substring(2, 2))
        RaiseEvent NewData(Channel, UpDown)
    End Sub

End Class