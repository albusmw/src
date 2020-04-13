Option Explicit On
Option Strict On

'''<summary>Class to change and add FITS file header information.</summary>
Public Class cFITSHeaderChanger

    '''<summary>Length of one header element.</summary>
    Const HeaderElementLength As Integer = 80
    '''<summary>Length of a header block - FITS files may contain an integer size of header blocks.</summary>
    Const HeaderBlockSize As Integer = 2880
    '''<summary>Number of header elements per header block.</summary>
    Public Shared ReadOnly HeaderElements As Integer = HeaderBlockSize \ HeaderElementLength

    '''<summary>Search a specific keyword in the passed list of header elements.</summary>
    '''<param name="HeaderElements">Header elements.</param>
    '''<param name="KeyWordToSearch">Keyword to search for.</param>
    Public Shared Function GetHeaderValue(ByRef HeaderElements As List(Of cFITSHeaderParser.sHeaderElement), ByVal KeyWordToSearch As String) As String
        For Each Entry As cFITSHeaderParser.sHeaderElement In HeaderElements
            If Entry.Keyword.Trim = KeyWordToSearch Then Return Entry.Value
        Next Entry
        Return String.Empty
    End Function

    Public Event Log(ByVal Text As String)

    Public ReadOnly Property HeaderElementsRead() As Integer
        Get
            Return MyHeaderElementsRead
        End Get
    End Property
    Private MyHeaderElementsRead As Integer = -1

    Public ReadOnly Property HeaderElementsChanged() As Integer
        Get
            Return MyHeaderElementsChanged
        End Get
    End Property
    Private MyHeaderElementsChanged As Integer = -1

    Public ReadOnly Property HeaderElementsInserted() As Integer
        Get
            Return MyHeaderElementsInserted
        End Get
    End Property
    Private MyHeaderElementsInserted As Integer = -1

    Public Shared Function Format(ByVal Value As Integer) As String
        Return Value.ToString().Replace(",", ".").Trim.PadLeft(20)
    End Function

    Public Shared Function Format(ByVal Value As DateTime) As String
        Return Microsoft.VisualBasic.Strings.Format(Value, "yyyy-MM-ddTHH:mm:ss").PadLeft(20)
    End Function

    Public Shared Function Format(ByVal Value As Double) As String
        If Value < 0 Then
            Return Value.ToString("0.0000000000000E+00").Replace(",", ".").Trim.PadLeft(18)
        Else
            Return "+" & Value.ToString("0.0000000000000E+00").Replace(",", ".").Trim.PadLeft(18)
        End If
    End Function

    Public Shared Function Format(ByVal Value As String) As String
        Return ("'" & Value & "'").PadRight(20)
    End Function

    Public Shared Function ReadHeader(ByVal File As String) As List(Of cFITSHeaderParser.sHeaderElement)

        Dim RetVal As New List(Of cFITSHeaderParser.sHeaderElement)

        If System.IO.File.Exists(File) = True Then

            'Open original file and create new list of header elements
            Dim FITS_stream As System.IO.FileStream = System.IO.File.OpenRead(File)

            Do

                'Get one header line which has 80 bytes
                Dim HeaderBytes(HeaderElementLength - 1) As Byte
                FITS_stream.Read(HeaderBytes, 0, HeaderBytes.Length)
                Dim SingleLine As String = System.Text.ASCIIEncoding.ASCII.GetString(HeaderBytes)

                'Exit on END detected
                If SingleLine.Trim.StartsWith("END") Then
                    Exit Do
                End If

                'Process only non-empty files
                If SingleLine.Trim.Length > 0 Then

                    'Get the keyword and the value
                    Dim HeaderElement As cFITSHeaderParser.sHeaderElement
                    HeaderElement.Keyword = SingleLine.Substring(0, 8)
                    HeaderElement.Value = SingleLine.Substring(9).Trim
                    HeaderElement.Comment = String.Empty
                    If HeaderElement.Value.Contains("/") Then
                        Dim SepPos As Integer = HeaderElement.Value.IndexOf("/")
                        HeaderElement.Comment = HeaderElement.Value.Substring(SepPos + 1).Trim
                        HeaderElement.Value = HeaderElement.Value.Substring(0, SepPos).Trim
                    End If
                    RetVal.Add(HeaderElement)

                End If

                'End on stream end
                If FITS_stream.Position = FITS_stream.Length Then Exit Do

            Loop Until 1 = 0

            FITS_stream.Close()

        End If

        Return RetVal

    End Function

    Public Function ChangeHeader(ByVal File As String, ByVal ValuesToChange As Dictionary(Of String, Object)) As Boolean
        Return ChangeHeader(File, File, ValuesToChange)
    End Function

    Public Function ChangeHeader(ByVal OriginalFile As String, ByVal NewFile As String, ByVal ValuesToChange As Dictionary(Of String, Object)) As Boolean

        'Detect in-place operation and correct
        Dim FileInPlace As Boolean = False
        If OriginalFile = NewFile Then
            NewFile = System.IO.Path.GetTempFileName
            FileInPlace = True
        End If

        'Open original file and create new list of header elements
        Dim FITS_stream As System.IO.FileStream = System.IO.File.OpenRead(OriginalFile)
        Dim NewHeaderElements As New List(Of String)

        'Read all header elements present
        MyHeaderElementsRead = 0
        MyHeaderElementsChanged = 0
        MyHeaderElementsInserted = 0

        Do

            'Get one header line which has 80 bytes
            MyHeaderElementsRead += 1
            Dim HeaderBytes(HeaderElementLength - 1) As Byte
            FITS_stream.Read(HeaderBytes, 0, HeaderBytes.Length)
            Dim SingleLine As String = System.Text.ASCIIEncoding.ASCII.GetString(HeaderBytes)

            'Get the keyword and the value
            Dim Keyword As String = SingleLine.Substring(0, 8)
            Dim Value As String = SingleLine.Substring(10)
            Dim Comment As String = String.Empty
            If Value.Contains("/") Then
                Dim SepPos As Integer = Value.IndexOf("/")
                Comment = Value.Substring(SepPos + 1)
                Value = Value.Substring(0, SepPos)
            End If

            'If the value to change is specified, change the value and mark as empty in the dictionary
            If ValuesToChange.ContainsKey(Keyword.Trim.ToUpper) Then
                MyHeaderElementsChanged += 1
                ChangeValue(Value, ValuesToChange(Keyword.Trim.ToUpper))
                ValuesToChange(Keyword.Trim.ToUpper) = Nothing
            End If

            'Exit on END detected
            If SingleLine.StartsWith("END") Then
                Exit Do
            Else
                'Log(SingleLine & "|")      -> Original element
                Dim NewElement As String = Keyword & "= " & Value & CStr(If(Comment.Length > 0, "/" & Comment, String.Empty))
                NewHeaderElements.Add(NewElement)
                RaiseEvent Log(NewElement)
            End If

        Loop Until 1 = 0

        'Inject header elements missing (that where not found in the dictionary yet)
        Dim HeaderElementsInserted As Integer = 0
        For Each Keyword As String In ValuesToChange.Keys
            If IsNothing(ValuesToChange(Keyword)) = False Then
                MyHeaderElementsInserted += 1
                Dim NewKeyword As String = Keyword.PadRight(8)
                Dim NewValue As String = Nothing
                Dim NewComment As String = String.Empty
                If IsArray(ValuesToChange(Keyword)) = True Then
                    NewValue = CStr(CType(ValuesToChange(Keyword), Array).GetValue(0))
                    NewComment = CStr(CType(ValuesToChange(Keyword), Array).GetValue(1))
                Else
                    NewValue = CStr(ValuesToChange(Keyword))
                    NewComment = String.Empty
                End If
                Dim NewElement As String = (NewKeyword & "= " & NewValue & CStr(If(NewComment.Length > 0, " /" & NewComment, String.Empty))).PadRight(HeaderElementLength)
                RaiseEvent Log(NewElement)
                NewHeaderElements.Add(NewElement)
                HeaderElementsInserted += 1
            End If
        Next Keyword

        'Inject empty lines to fill the required header block size of 
        Dim TotalHeaderLine As Integer = HeaderElementsRead + HeaderElementsInserted
        Do
            If TotalHeaderLine Mod HeaderElements = 0 Then Exit Do
            Dim NewElement As String = New String(Chr(&H20), HeaderElementLength)
            NewHeaderElements.Add(NewElement)
            TotalHeaderLine += 1
        Loop Until 1 = 0

        'Inject END
        Dim EndElement As String = "END".PadRight(HeaderElementLength)
        NewHeaderElements.Add(EndElement)
        RaiseEvent Log(EndElement)

        'Check length
        For Idx As Integer = 0 To NewHeaderElements.Count - 1
            If NewHeaderElements(Idx).Length > HeaderElementLength Then
                NewHeaderElements(Idx) = NewHeaderElements(Idx).Substring(0, HeaderElementLength)
            End If
        Next Idx

        'Write the new header to the new file
        System.IO.File.WriteAllBytes(NewFile, System.Text.ASCIIEncoding.ASCII.GetBytes(Join(NewHeaderElements.ToArray, "")))

        'Close original stream
        FITS_stream.Close()

        'Position to 1st binary element
        Dim SeekStart As Integer = CInt(Math.Ceiling((HeaderElementsRead / HeaderElements))) * HeaderElements * HeaderElementLength

        'Copy original stream data to new FITS file
        AppendBinaryContent(OriginalFile, SeekStart, NewFile)

        'If in-place, copy and delete
        If FileInPlace Then
            System.IO.File.Copy(NewFile, OriginalFile, True)
            System.IO.File.Delete(NewFile)
        End If

        Return True

    End Function

    Private Sub AppendBinaryContent(ByVal SourceFile As String, ByVal SourceFileStart As Integer, ByVal DestinationFile As String)

        Dim InStream As New System.IO.FileStream(SourceFile, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None, 1024, IO.FileOptions.Asynchronous)
        Dim CopyStream(32767) As Byte
        Dim OutStream As New System.IO.FileStream(DestinationFile, IO.FileMode.Append, IO.FileAccess.Write, IO.FileShare.None, 1024, IO.FileOptions.Asynchronous)

        InStream.Seek(SourceFileStart, IO.SeekOrigin.Begin)

        Do
            Dim BytesRead As Integer = InStream.Read(CopyStream, 0, CopyStream.Length)
            If BytesRead = 0 Then Exit Do
            OutStream.Write(CopyStream, 0, BytesRead)

        Loop Until 1 = 0

        OutStream.Flush()
        InStream.Close()
        OutStream.Close()

    End Sub

    Private Sub ChangeValue(ByRef ValueToChange As String, ByVal NewValue As Object)

        Dim NewValueString As String = String.Empty
        If IsArray(NewValue) = False Then
            NewValueString = CStr(NewValue)
        Else
            NewValueString = CStr(CType(NewValue, Array).GetValue(0))
        End If

        Dim OldLength As Integer = ValueToChange.Trim.Length
        Dim NewLength As Integer = NewValueString.Trim.Length

        'Count right spaces
        Dim RightIdx As Integer = ValueToChange.Length - 1
        Dim RightSpaces As Integer = 0
        Do
            If ValueToChange.Substring(RightIdx, 1) = " " Then
                RightSpaces += 1
                RightIdx -= 1
            Else
                Exit Do
            End If
        Loop Until RightIdx = 0

        'Count left spaces
        Dim LeftIdx As Integer = 0
        Dim LeftSpaces As Integer = 0
        Do
            If ValueToChange.Substring(LeftIdx, 1) = " " Then
                LeftSpaces += 1
                LeftIdx += 1
            Else
                Exit Do
            End If
        Loop Until LeftIdx = ValueToChange.Length - 1

        'Entry STARTS with spaces
        If RightSpaces = 0 Or RightSpaces = 1 Then
            If NewLength >= OldLength Then
                LeftSpaces -= (NewLength - OldLength) : If LeftSpaces < 0 Then LeftSpaces = 0
            Else
                LeftSpaces += (OldLength - NewLength)
            End If
        End If

        'Entry ENDS with spaces
        If LeftSpaces = 0 Or LeftSpaces = 1 Then
            If NewLength >= OldLength Then
                RightSpaces -= (NewLength - OldLength) : If RightSpaces < 0 Then RightSpaces = 0
            Else
                RightSpaces += (OldLength - NewLength)
            End If
        End If

        'Fall-back
        ValueToChange = New String(CChar(" "), LeftSpaces) & NewValueString & New String(CChar(" "), RightSpaces)

    End Sub

End Class
