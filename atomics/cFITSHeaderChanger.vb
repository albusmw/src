Option Explicit On
Option Strict On

'''<summary>Class to change and add FITS file header information.</summary>
Public Class cFITSHeaderChanger

    '''<summary>List of all keywords that where not recognized during parsing.</summary>
    Public Shared UnknownKeywords As New List(Of String)

    '''<summary>Search a specific keyword in the passed list of header elements.</summary>
    '''<param name="HeaderElements">Header elements.</param>
    '''<param name="KeyWordToSearch">Keyword to search for.</param>
    Public Shared Function GetHeaderValue(ByRef HeaderElements As List(Of cFITSHeaderParser.sHeaderElement), ByVal KeyWordToSearch As eFITSKeywords) As Object
        For Each Entry As cFITSHeaderParser.sHeaderElement In HeaderElements
            If Entry.Keyword = KeyWordToSearch Then Return Entry.Value
        Next Entry
        Return String.Empty
    End Function

    '''<summary>Number of block to read in to get the complete header - a "good guess".</summary>
    Public Shared Property ReadHeaderBlockCount As Integer = 3

    '''<summary>Number of block to read in to get the complete header.</summary>
    Public Shared ReadOnly Property ReadHeaderByteCount() As Integer = ReadHeaderBlockCount * FITSSpec.HeaderBlockSize

    '''<summary>Parse the header of the given file.</summary>
    '''<param name="FITSFile">FITS file to parse</param>
    Public Shared Function ParseHeader(ByVal FITSFile As String) As List(Of cFITSHeaderParser.sHeaderElement)
        Dim Dummy As Integer = -1
        Return ParseHeader(FITSFile, Dummy)
    End Function

    '''<summary>Parse the header of the given file.</summary>
    '''<param name="FITSFile">FITS file to parse</param>
    Public Shared Function ParseHeader(ByVal FITSFile As String, ByRef DataStartPos As Integer) As List(Of cFITSHeaderParser.sHeaderElement)
        Dim RetNothing As New List(Of cFITSHeaderParser.sHeaderElement)
        If System.IO.File.Exists(FITSFile) = True Then
            Dim HeaderBytes(ReadHeaderByteCount - 1) As Byte
            System.IO.File.OpenRead(FITSFile).Read(HeaderBytes, 0, HeaderBytes.Length)
            Return ParseHeader(HeaderBytes, DataStartPos)
        End If
        Return RetNothing
    End Function

    '''<summary>Parse the given header.</summary>
    '''<param name="HeaderBytes">Header bytes read from the file.</param>
    '''<param name="DataStartPos">Return value - data start position relative to file start.</param>
    '''<returns>List of header elements.</returns>
    Public Shared Function ParseHeader(ByVal HeaderBytes As Byte(), ByRef DataStartPos As Integer) As List(Of cFITSHeaderParser.sHeaderElement)

        Dim RetVal As New List(Of cFITSHeaderParser.sHeaderElement)
        Dim BytesRead As Integer = 0
        Dim FormatProvider As IFormatProvider = Globalization.CultureInfo.InvariantCulture
        Dim EndFound As Boolean = False

        'Check if this is a FITS file
        Dim MagicString As String = GetString(HeaderBytes, 0, 6)
        If MagicString <> "SIMPLE" Then Return New List(Of cFITSHeaderParser.sHeaderElement)

        Dim HeaderPtr As Integer = 0
        Do

            'Get one header line which has 80 bytes
            Dim SingleLine As String = GetString(HeaderBytes, HeaderPtr, FITSSpec.HeaderElementLength)
            HeaderPtr += FITSSpec.HeaderElementLength
            BytesRead += FITSSpec.HeaderElementLength

            'Exit on END detected
            If SingleLine.Trim.StartsWith("END") Then
                EndFound = True
                Exit Do
            End If

            'Process only non-empty files
            If SingleLine.Trim.Length > 0 Then

                'Get keyword, value and comment
                Dim HeaderElement As cFITSHeaderParser.sHeaderElement
                Dim KeywordString As String = SingleLine.Substring(0, FITSSpec.HeaderKeywordLength)
                HeaderElement.Keyword = cFITSHeaderParser.GetKeywordEnum(KeywordString)
                If HeaderElement.Keyword = eFITSKeywords.UNKNOWN Then
                    If Not UnknownKeywords.Contains(KeywordString) Then UnknownKeywords.Add(KeywordString)
                End If
                Dim Value As String = SingleLine.Substring(FITSSpec.HeaderKeywordLength + 1).Trim               'we take +1 as the space after the "=" sign is sometimes not added / filled with data
                HeaderElement.Value = Value
                HeaderElement.Comment = String.Empty
                If CStr(Value).Contains("/") Then
                    Dim SepPos As Integer = CStr(Value).IndexOf("/")
                    HeaderElement.Comment = CStr(Value).Substring(SepPos + 1).Trim
                    Value = CStr(Value).Substring(0, SepPos).Trim
                End If
                'Try to auto-detect the value
                If Value.StartsWith("'") And Value.EndsWith("'") Then
                    'String
                    If Value.Length > 2 Then HeaderElement.Value = Value.Substring(1, Value.Length - 2)
                Else
                    Dim ValueAsInt As Integer = Integer.MinValue
                    Dim ValueAsDouble As Double = Double.NaN
                    If Integer.TryParse(Value, ValueAsInt) = True Then
                        HeaderElement.Value = ValueAsInt
                    Else
                        If Double.TryParse(Value, Globalization.NumberStyles.Float, FormatProvider, ValueAsDouble) = True Then
                            HeaderElement.Value = ValueAsDouble
                        End If
                    End If
                End If

                'Store final element as new header element
                RetVal.Add(HeaderElement)

            End If

            'End on stream end
            If HeaderPtr = HeaderBytes.Length Then Exit Do

        Loop Until 1 = 0

        If EndFound = True Then
            DataStartPos = CInt(Math.Ceiling(BytesRead / FITSSpec.HeaderBlockSize) * FITSSpec.HeaderBlockSize)
        Else
            DataStartPos = -1                                                                   'END was not detected -> data start unknown ...
        End If
        Return RetVal

    End Function

    '''<summary>Change the header of the given file in-place.</summary>
    '''<param name="File">File to change header.</param>
    '''<param name="ValuesToChange">Keyword-value dictionary of elements to change.</param>
    '''<returns>Empty string for success; error codes else.</returns>
    Public Shared Function ChangeHeader(ByVal File As String, ByVal ValuesToChange As Dictionary(Of String, Object)) As String
        Return ChangeHeader(File, File, ValuesToChange)
    End Function

    '''<summary>Change the header of the given file.</summary>
    '''<param name="OriginalFile">File to get data from.</param>
    '''<param name="NewFile">New file to create.</param>
    '''<param name="ValuesToChange">Keyword-value dictionary of elements to change - dictionary value can be a 2-element array with {value,comment}.</param>
    '''<returns>Empty string for success; error codes else.</returns>
    Public Shared Function ChangeHeader(ByVal OriginalFile As String, ByVal NewFile As String, ByVal ValuesToChange As Dictionary(Of String, Object)) As String

        Dim FITS_stream As System.IO.FileStream

        'Detect in-place operation and correct
        Dim FileInPlace As Boolean = False
        If OriginalFile = NewFile Then
            NewFile = System.IO.Path.GetTempFileName
            FileInPlace = True
        End If

        '======================================================================================================
        'Read header
        '------------------------------------------------------------------------------------------------------

        'Read all header elements present as-is
        FITS_stream = System.IO.File.OpenRead(OriginalFile)
        Dim AllCards As New List(Of Byte())
        Do
            'Get one header line which has 80 bytes
            Dim CardBytes(FITSSpec.HeaderElementLength - 1) As Byte
            FITS_stream.Read(CardBytes, 0, CardBytes.Length)
            'Exit on END detected
            If System.Text.ASCIIEncoding.ASCII.GetString(CardBytes, 0, 3) = "END" Then
                Exit Do
            Else
                AllCards.Add(CardBytes)
            End If
        Loop Until 1 = 0
        FITS_stream.Close()

        'Calculate how many blocks are occupied
        Dim BlocksUsed As Integer = CInt(Math.Ceiling(AllCards.Count / FITSSpec.HeaderElements))

        '======================================================================================================
        'Process changes
        '------------------------------------------------------------------------------------------------------

        'Move over all elements and change present entries
        Dim ChangedValues As New List(Of String)
        For Idx As Integer = 0 To AllCards.Count - 1
            Dim Keyword As String = GetString(AllCards(Idx), 0, FITSSpec.HeaderKeywordLength).Trim
            If ValuesToChange.ContainsKey(Keyword) Then
                ChangeValueInCard(AllCards(Idx), ValuesToChange(Keyword))
                ChangedValues.Add(Keyword)
            End If
        Next Idx

        'Inject header elements missing (that where not found in the dictionary yet)
        For Each Keyword As String In ValuesToChange.Keys
            If IsNothing(ValuesToChange(Keyword)) = False Then
                If ChangedValues.Contains(Keyword) = False Then
                    Dim NewKeyword As String = PadKeyword(Keyword)
                    Dim NewValue As String = Nothing
                    Dim NewComment As String = String.Empty
                    If IsArray(ValuesToChange(Keyword)) = True Then
                        NewValue = CStr(CType(ValuesToChange(Keyword), Array).GetValue(0))
                        NewComment = CStr(CType(ValuesToChange(Keyword), Array).GetValue(1))
                    Else
                        NewValue = CStr(ValuesToChange(Keyword))
                        NewComment = String.Empty
                    End If
                    'Create the complete new element and trim or align to correct length (HeaderElementLength=80)
                    Dim NewElement As String = (NewKeyword & FITSSpec.HeaderEqualString & PadValue(NewValue) & CStr(If(NewComment.Length > 0, " /" & NewComment, String.Empty)))
                    AllCards.Add(GetBytes(FITSSpec.EnsureCorrectLength(NewElement)))
                End If
            End If
        Next Keyword

        'Inject END and empty lines to fill the required header block size of 
        AllCards.Add(GetBytes("END".PadRight(FITSSpec.HeaderElementLength)))
        Dim EmptyCard As String = New String(Chr(&H20), FITSSpec.HeaderElementLength)
        Do
            If AllCards.Count Mod FITSSpec.HeaderElements = 0 Then Exit Do
            AllCards.Add(GetBytes(EmptyCard))
        Loop Until 1 = 0

        'Get complete buffer as byte
        Dim HeaderBytes As New List(Of Byte)
        For Each Card As Byte() In AllCards
            HeaderBytes.AddRange(Card)
        Next Card

        'Calculate how many blocks are occupied
        Dim BlocksRequired As Integer = CInt(Math.Ceiling(AllCards.Count / FITSSpec.HeaderElements))

        '======================================================================================================
        'Adjust output file
        '------------------------------------------------------------------------------------------------------

        'For inplace and no header size change, just overwrite old header
        If (BlocksUsed = BlocksRequired) And (FileInPlace = True) Then
            FITS_stream = System.IO.File.OpenWrite(OriginalFile)
            FITS_stream.Write(HeaderBytes.ToArray, 0, HeaderBytes.Count)
            FITS_stream.Close()
        End If


        ''Write the new header to the new file
        'System.IO.File.WriteAllBytes(NewFile, System.Text.ASCIIEncoding.ASCII.GetBytes(Join(NewHeaderElements.ToArray, "")))

        ''Close original stream
        'FITS_stream.Close()

        ''Position to 1st binary element
        'Dim SeekStart As Integer = CInt(Math.Ceiling((HeaderElementsRead / HeaderElements))) * HeaderElements * HeaderElementLength

        ''Copy original stream data to new FITS file
        'AppendBinaryContent(OriginalFile, SeekStart, NewFile)

        ''If in-place, copy and delete
        'If FileInPlace Then
        '    System.IO.File.Copy(NewFile, OriginalFile, True)
        '    System.IO.File.Delete(NewFile)
        'End If

        Return String.Empty

    End Function

    '''<summary>Change the value in the given card.</summary>
    '''<param name="Card">Card.</param>
    '''<param name="NewValue">New value to set.</param>
    Private Shared Sub ChangeValueInCard(ByRef Card As Byte(), ByVal NewValue As Object)

        'Get value and comment (if present) of the given card byte sequence
        Dim OldValueString As String = GetString(Card, 10, Card.Length - 11)
        Dim Comment As String = String.Empty
        If OldValueString.Contains("/") Then
            Dim SepPos As Integer = OldValueString.IndexOf("/")
            Comment = OldValueString.Substring(SepPos + 1)              'comment contains the "/"
            OldValueString = OldValueString.Substring(0, SepPos)
        End If
        OldValueString = OldValueString.Trim(" "c)                      'remove spaces

        'Get the string representation of the given value - for {value,comment} take value only
        Dim NewValueString As String = String.Empty
        If IsArray(NewValue) = False Then
            NewValueString = PadValue(cFITSType.AsString(NewValue))
        Else
            NewValueString = PadValue(cFITSType.AsString(CType(NewValue, Array).GetValue(0)))
            Comment = " / " & CStr(CType(NewValue, Array).GetValue(1))  'we use space before and after as this is more readable ...
        End If

        'TODO:
        'Changed values are always right-aligned; the alignment of the original file is not taken into account!

        'Form new card byte array
        Dim NewCard As New List(Of Byte)
        NewCard.AddRange(Card.Take(8))
        NewCard.AddRange(GetBytes(FITSSpec.HeaderEqualString))
        NewCard.AddRange(GetBytes(NewValueString))
        NewCard.AddRange(GetBytes(Comment))
        Card = FITSSpec.EnsureCorrectLength(NewCard).ToArray

    End Sub

    '''<summary>Convert string to bytes.</summary>
    '''<remarks>Done in a central position to ensure same encoding.</remarks>
    Private Shared Function GetBytes(ByVal Text As String) As Byte()
        Return System.Text.ASCIIEncoding.ASCII.GetBytes(Text)
    End Function

    '''<summary>Convert bytes to string.</summary>
    '''<remarks>Done in a central position to ensure same encoding.</remarks>
    Private Shared Function GetString(ByVal Bytes As Byte(), ByVal Index As Integer, ByVal Count As Integer) As String
        Return System.Text.ASCIIEncoding.ASCII.GetString(Bytes, Index, Count)
    End Function

    '''<summary>Pad the given string.</summary>
    '''<remarks>Done in a central position to ensure same behaviour.</remarks>
    Private Shared Function PadKeyword(ByVal Text As String) As String
        Return Text.PadRight(FITSSpec.HeaderKeywordLength)
    End Function

    '''<summary>Pad the given string.</summary>
    '''<remarks>Done in a central position to ensure same behaviour.</remarks>
    Private Shared Function PadValue(ByVal Text As String) As String
        Return Text.PadLeft(FITSSpec.HeaderValueLength)
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
