'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    'TODO:
    '- If list is not sorted, multiple sections can appear -> collect all sections on top ...
    '- Entries with no section (starting in line 1 until a section starts) must stay on top also witout sorting ...

    '''<summary>Manage input and output for INI classes.</summary>
    Public Class cINI_IO

        '''<summary>Indicate if some valid content was loaded from a INI file.</summary>
        Public ReadOnly Property Loaded() As Boolean
            Get
                If IsNothing(Entries) = True Then Return False
                If Entries.Count = 0 Then Return False
                Return True
            End Get
        End Property

        Private Entries As New Dictionary(Of String, List(Of String))

        '''<summary>Allow multiple values for 1 key.</summary>
        Public Property AllowMultipleEntries() As Boolean
            Get
                Return MyAllowMultipleEntries
            End Get
            Set(ByVal value As Boolean)
                MyAllowMultipleEntries = value
            End Set
        End Property
        Private MyAllowMultipleEntries As Boolean = False

        '''<summary>Sort all entries before storing.</summary>
        Public Property SortBeforeStore() As Boolean
            Get
                Return MySortBeforeStore
            End Get
            Set(ByVal value As Boolean)
                MySortBeforeStore = value
            End Set
        End Property
        Private MySortBeforeStore As Boolean = True

        '''<summary>Sign(s) that are used to indicate a comment.</summary>
        Public Property CommentSigns() As List(Of String)
            Get
                Return MyCommentSigns
            End Get
            Set(ByVal value As List(Of String))
                MyCommentSigns = value
            End Set
        End Property
        Private MyCommentSigns As New List(Of String)(New String() {";", "%", "/"})

        Private Function ContentValid() As Boolean
            If IsNothing(Entries) = True Then Return False
            If Entries.Count = 0 Then Return False
            Return True
        End Function

        '''<summary>Load the content of the given INI file.</summary>
        '''<remarks>All previous loaded entries will be removed.</remarks>
        Public Sub Load(ByVal FileName As String)
            If System.IO.File.Exists(FileName) = True Then
                Entries = New Dictionary(Of String, List(Of String))
                ParseContent(System.IO.File.ReadAllLines(FileName))
            End If
        End Sub

        '''<summary>Load the content of the given string array.</summary>
        '''<remarks>All previous loaded entries will be removed.</remarks>
        Public Sub Load(ByVal Content As String())
            If IsNothing(Content) = True Then Exit Sub
            If Content.Length = 0 Then Exit Sub
            Entries = New Dictionary(Of String, List(Of String))
            ParseContent(Content)
        End Sub

        '''<summary>Save the content of the given INI file.</summary>
        Public Sub Save(ByVal FileName As String)
            'Create a flat list of elements to store
            Dim ToStore As New List(Of String)
            For Each Key As String In Entries.Keys
                For Each Value As String In Entries(Key)
                    ToStore.Add(Key & "=" & Value)
                Next Value
            Next Key
            'Sort if requested
            If SortBeforeStore = True Then ToStore.Sort(AddressOf INISorter)
            'Parse through the list
            Dim FileContent As New List(Of String)
            Dim CurrentSection As String = String.Empty
            For Each Entry As String In ToStore
                If Entry.StartsWith("[") And Entry.Contains("]") Then
                    'Get the current section and start a new one if section does not match previous section
                    Dim NewSection As String = Entry.Substring(1, Entry.IndexOf("]") - 1)
                    If NewSection <> CurrentSection Then
                        CurrentSection = NewSection
                        FileContent.Add("[" & NewSection & "]")
                    End If
                    Dim KeyValue As String = Entry.Substring(Entry.IndexOf("]") + 2)
                    FileContent.Add(KeyValue)
                Else
                    FileContent.Add(Entry)
                End If
            Next Entry
            System.IO.File.WriteAllLines(FileName, FileContent.ToArray)
        End Sub

        '''<summary>Ensure that elements without "[" end in the beginning.</summary>
        Private Function INISorter(ByVal X As String, ByVal Y As String) As Integer
            If X.StartsWith("[") = True And Y.StartsWith("[") = False Then Return 1
            If X.StartsWith("[") = False And Y.StartsWith("[") = True Then Return -1
            Return X.CompareTo(Y)
        End Function

        '''<summary>Parse the passed content.</summary>
        Private Sub ParseContent(ByRef Content As String())
            Dim CurrentSection As String = String.Empty
            For Each Line As String In Content
                Dim LineTrim As String = Line.Trim
                'Set the new section if there is any
                If LineTrim.StartsWith("[") And LineTrim.EndsWith("]") Then
                    CurrentSection = LineTrim.Substring(1, LineTrim.Length - 2)
                Else
                    If LineTrim.Contains("=") = False Then
                        'Line is not a section and also does not contain values -> not a valid line ...
                    Else
                        If CommentSigns.Contains(LineTrim.Substring(0, 1)) Then
                            'Comment
                        Else
                            'Compose key
                            Dim Key As String = LineTrim.Substring(0, LineTrim.IndexOf("="))
                            If String.IsNullOrEmpty(CurrentSection) = False Then Key = "[" & CurrentSection & "]." & Key
                            'Get value
                            Dim Value As String = LineTrim.Substring(LineTrim.IndexOf("=") + 1)
                            'Create a new entry if the key does not exist
                            If Entries.ContainsKey(Key) = False Then
                                Entries.Add(Key, New List(Of String)(New String() {Value}))
                            Else
                                If AllowMultipleEntries Then
                                    Entries(Key).Add(Value)
                                Else
                                    Entries(Key) = New List(Of String)(New String() {Value})
                                End If
                            End If
                        End If
                    End If
                End If
            Next Line
        End Sub

        '================================================================================

        ''' <summary>Get the value of the INI file entry specified.</summary>
        Public Function [Get](ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As Boolean) As Boolean
            Dim RawRead As String = [Get](Section, KeyName, String.Empty)
            If String.IsNullOrEmpty(RawRead) Then
                [Set](Section, KeyName, CStr(IIf(DefaultValue = True, "TRUE", "FALSE")).Trim)                   'key does not exist -> set
                Return DefaultValue
            Else
                Select Case RawRead.ToUpper
                    Case "1", "TRUE", "YES"
                        Return True
                    Case "0", "FALSE", "NO"
                        Return False
                    Case Else
                        Return False
                End Select
            End If
        End Function

        ''' <summary>Get the value of the INI file entry specified.</summary>
        Public Function [Get](ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As Integer) As Integer
            Dim RawRead As String = [Get](Section, KeyName, String.Empty)
            If String.IsNullOrEmpty(RawRead) Then
                [Set](Section, KeyName, CStr(DefaultValue).Trim)                   'key does not exist -> set
                Return DefaultValue
            Else
                Dim ParsedValued As Integer = 0
                Dim CouldRead As Boolean = Integer.TryParse(RawRead, ParsedValued)
                If CouldRead = True Then
                    Return ParsedValued
                Else
                    [Set](Section, KeyName, CStr(DefaultValue).Trim)                   'key does not exist -> set
                    Return DefaultValue
                End If
            End If
        End Function

        ''' <summary>Get the value of the INI file entry specified.</summary>
        Public Function [Get](ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As Double) As Double
            Dim RawRead As String = [Get](Section, KeyName, String.Empty)
            If String.IsNullOrEmpty(RawRead) Then
                [Set](Section, KeyName, CStr(DefaultValue).Trim.Replace(",", "."))                   'key does not exist -> set
                Return DefaultValue
            Else
                Dim ParsedValued As Double = 0
                Try
                    Return Val(RawRead.Replace(",", "."))
                Catch ex As Exception
                    [Set](Section, KeyName, CStr(DefaultValue).Trim.Replace(",", "."))                   'key does not exist -> set
                    Return DefaultValue
                End Try
            End If
        End Function

        ''' <summary>Get the value of the INI file entry specified.</summary>
        Public Function [Get](ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
            If ContentValid() = True Then
                Dim Key As String = String.Empty
                If String.IsNullOrEmpty(Section) = False Then Key = "[" & Section & "]." & KeyName Else Key = KeyName
                If Entries.ContainsKey(Key) Then
                    Dim RetVal As List(Of String) = Entries(Key)
                    If IsNothing(RetVal) = True Then Return String.Empty
                    Select Case RetVal.Count
                        Case 0 : Return String.Empty
                        Case 1 : Return RetVal(0)
                        Case Else : Return Join(RetVal.ToArray, "|")
                    End Select
                End If
            End If
            [Set](Section, KeyName, DefaultValue)                   'key does not exist -> set
            Return DefaultValue
        End Function

        ''' <summary>Set the value of the INI file entry specified.</summary>
        Public Sub [Set](ByVal Section As String, ByVal KeyName As String, ByVal Value As String)
            'Generate key to set
            Dim Key As String = KeyName
            If String.IsNullOrEmpty(Section) = False Then Key = "[" & Section & "]." & Key
            'If there is no content in the entries, create new lust
            If ContentValid() = False Then
                Entries = New Dictionary(Of String, List(Of String))
                Entries.Add(Key, New List(Of String)(New String() {Value}))
            Else
                If Entries.ContainsKey(Key) = False Then
                    Entries.Add(Key, New List(Of String)(New String() {Value}))
                Else
                    Entries(Key) = New List(Of String)(New String() {Value})
                End If
            End If
        End Sub

    End Class

End Namespace