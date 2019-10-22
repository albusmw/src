Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Class to write CSV files.</summary>
    '''<remarks>The class takes the configured decimal separator from the system configuration.</remarks>
    Public Class cCSVBuilder

        Private Rows As List(Of Dictionary(Of String, String))
        Private CurrentRow As Integer = -1
        Private CSVSplitSign As String = ";"
        Private ColumnSplitter As String = "|"

        Public Sub New()
            Clear()
        End Sub

        Public Shared Sub JustDump(ByVal FileName As String, ByRef Vector() As Double)
            Dim MyDumper As New cCSVBuilder
            For Idx As Integer = 0 To Vector.GetUpperBound(0)
                MyDumper.StartRow()
                MyDumper.AddColumnValue("Idx", Idx)
                MyDumper.AddColumnValue("Value", Vector(Idx))
            Next Idx
            System.IO.File.WriteAllText(FileName, MyDumper.CreateCSV())
        End Sub

        Public Shared Sub JustDump(ByVal FileName As String, ByRef Vector1() As Double, ByRef Vector2() As Double)
            Dim MyDumper As New cCSVBuilder
            For Idx As Integer = 0 To Vector1.GetUpperBound(0)
                MyDumper.StartRow()
                MyDumper.AddColumnValue("Idx", Idx)
                MyDumper.AddColumnValue("Value1", Vector1(Idx))
                MyDumper.AddColumnValue("Value2", Vector2(Idx))
            Next Idx
            System.IO.File.WriteAllText(FileName, MyDumper.CreateCSV())
        End Sub

        '''<summary>Reset the CSV class.</summary>
        Public Sub Clear()
            Rows = New List(Of Dictionary(Of String, String))
            CurrentRow = -1
        End Sub

        '''<summary>Start with writing to a new row.</summary>
        Public Sub StartRow()
            Rows.Add(New Dictionary(Of String, String))
            CurrentRow += 1
        End Sub

        '''<summary>Add a new double value.</summary>
        '''<param name="ColumnName">Name of the column to be added.</param>
        '''<param name="Value">Value to be added</param>
        Public Sub AddColumnValue(ByVal ColumnName As String, ByVal Value As Double)
            AddColumnValue(ColumnName, FormaterValueToWrite(Value))
        End Sub

        '''<summary>Add a new double value.</summary>
        '''<param name="ColumnName">Name of the column to be added.</param>
        '''<param name="Value">Value to be added</param>
        '''<param name="FormatString">Format to use.</param>
        Public Sub AddColumnValue(ByVal ColumnName As String, ByVal Value As Double, ByVal FormatString As String)
            AddColumnValue(ColumnName, FormaterValueToWrite(Value, FormatString))
        End Sub

        '''<summary>Add a new boolean value.</summary>
        '''<param name="ColumnName">Name of the column to be added.</param>
        '''<param name="Value">Value to be added</param>
        Public Sub AddColumnValue(ByVal ColumnName As String, ByVal Value As Boolean)
            AddColumnValue(ColumnName, FormaterValueToWrite(Value))
        End Sub

        '''<summary>Add a new value.</summary>
        '''<param name="ColumnName">Name of the column to be added.</param>
        '''<param name="Value">Value to be added</param>
        Public Sub AddColumnValue(ByVal ColumnName As String, ByVal Value As String)
            If CurrentRow > -1 Then
                With Rows(CurrentRow)
                    If .ContainsKey(ColumnName) = True Then
                        .Item(ColumnName) = Value
                    Else
                        .Add(ColumnName, Value)
                    End If
                End With
            End If
        End Sub

        '''<summary>Get the last row (e.g. to display this results in a text box).</summary>
        Public Function GetLastLine() As String
            Dim Values As New List(Of String)
            For Each Key As String In Rows(Rows.Count - 1).Keys
                Values.Add(Rows(Rows.Count - 1)(Key))
            Next Key
            Return Join(Values.ToArray, ColumnSplitter)
        End Function

        Private Sub GetLastEntries(ByRef Header As List(Of String), ByRef Values As List(Of String))
            GetLastEntries(Header, Values, True)
        End Sub

        '''<summary>Get the last line of the CSV file, including the headers.</summary>
        '''<param name="Header">List of header elements.</param>
        '''<param name="Values">List of entries.</param>
        '''<param name="UseLastRowOnly">If true, only the last line is used - use if the number of columns is fixed from the first line on.</param>
        Private Sub GetLastEntries(ByRef Header As List(Of String), ByRef Values As List(Of String), ByVal UseLastRowOnly As Boolean)

            Header = New List(Of String)
            Values = New List(Of String)

            Dim FirstIdx As Integer = CInt(IIf(UseLastRowOnly = True, Rows.Count - 1, 0))

            Dim Columns As New Dictionary(Of String, List(Of String))
            For Idx As Integer = FirstIdx To Rows.Count - 1
                For Each Key As String In Rows(Idx).Keys
                    If Columns.ContainsKey(Key) = False Then
                        Columns.Add(Key, New List(Of String))
                        For EmptyEntry As Integer = 0 To Idx - 1
                            Columns(Key).Add("--")
                        Next EmptyEntry
                    End If
                    Columns(Key).Add(Rows(Idx)(Key))
                Next Key
            Next Idx

            'Create the content to the CSV file (header is already there)
            For Each Column As String In Columns.Keys
                Header.Add(Column)
                Values.Add(Columns(Column).Item(Columns(Column).Count - 1))
            Next Column

            For Idx As Integer = 0 To Header.Count - 1
                If Header(Idx).Length > Values(Idx).Length Then
                    Values(Idx) = Space(Header(Idx).Length - Values(Idx).Length) & Values(Idx)
                Else
                    Header(Idx) = Space(Values(Idx).Length - Header(Idx).Length) & Header(Idx)
                End If
            Next Idx

        End Sub

        '''<summary>Get the header and the last row (e.g. to display this results in a text box).</summary>
        Public Function GetLastLineWithHeader() As String()

            Dim Header As New List(Of String)
            Dim Values As New List(Of String)
            GetLastEntries(Header, Values)

            Dim RetVal(1) As String
            RetVal(0) = Join(Header.ToArray, ColumnSplitter)
            RetVal(1) = Join(Values.ToArray, ColumnSplitter)
            Return RetVal

        End Function

        '''<summary>Get header and last line as dictionary.</summary>
        Public Function GetCurrentLineAsDictionary() As Dictionary(Of String, String)

            Dim Header As New List(Of String)
            Dim Value As New List(Of String)
            GetLastEntries(Header, Value)

            Dim RetVal As New Dictionary(Of String, String)
            For Idx As Integer = 0 To Header.Count - 1
                RetVal.Add(Header(Idx), Value(Idx))
            Next Idx
            Return RetVal

        End Function

        '''<summary>Create the complete content of the CSV file.</summary>
        '''<returns>CSV file content.</returns>
        Public Function CreateCSV() As String
            Return CreateCSV(False)
        End Function

        '''<summary>Create the complete content of the CSV file or only the last line.</summary>
        '''<param name="LastLineOnly">TRUE to get last line only, FALSE else.</param>
        '''<returns>CSV file content.</returns>
        Public Function CreateCSV(ByVal LastLineOnly As Boolean) As String

            Dim Columns As New Dictionary(Of String, List(Of String))

            For Idx As Integer = 0 To Rows.Count - 1

                For Each Key As String In Rows(Idx).Keys

                    If Columns.ContainsKey(Key) = False Then
                        Columns.Add(Key, New List(Of String))
                        For EmptyEntry As Integer = 0 To Idx - 1
                            Columns(Key).Add("--")
                        Next EmptyEntry
                    End If

                    Columns(Key).Add(Rows(Idx)(Key))

                Next Key

            Next Idx

            'Create the content to the CSV file (header is already there)
            Dim Lines As String() = {} : Array.Resize(Lines, Rows.Count + 1)
            For Each Column As String In Columns.Keys
                Lines(0) &= Column & CSVSplitSign
                For Idx As Integer = 0 To Columns(Column).Count - 1
                    Lines(Idx + 1) &= Columns(Column).Item(Idx) & CSVSplitSign
                Next Idx
            Next Column

            'Remove last ";"
            For Idx As Integer = 0 To Lines.GetUpperBound(0)
                If IsNothing(Lines(Idx)) = False Then
                    If Lines(Idx).Length > 0 Then
                        Lines(Idx) = Lines(Idx).Substring(0, Lines(Idx).Length - 1)
                    End If
                End If
            Next Idx

            If LastLineOnly = False Then
                Return Join(Lines, Environment.NewLine)
            Else
                Return Lines(Lines.GetUpperBound(0))
            End If

        End Function

        Private Shared Function FormaterValueToWrite(Of T)(ByVal Value As T) As String
            Return Str(Value).Trim.Replace(".", Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
        End Function

        Private Shared Function FormaterValueToWrite(ByVal Value As Double, ByVal FormatString As String) As String
            Return Format(Value, FormatString).Trim.Replace(".", Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
        End Function

    End Class

    '''<summary>Class for handling CSV (comma separated file) access.</summary>
    Public Class cCSVWriter

        '''<summary>One single row.</summary>
        Private Structure sRow
            Public Cells() As String
            Public Sub New(ByRef Cells() As String)
                Me.Cells = Cells
            End Sub
        End Structure

        Private Rows As List(Of sRow)

        '''<summary>Initilize by create a new sRow list instance.</summary>
        Public Sub New()
            Rows = New List(Of sRow)
        End Sub

        '''<summary>Add a list of cells to the content.</summary>
        '''<param name="Cells">List of string values to add.</param>
        '''<remarks></remarks>
        Public Sub Append(ByVal Cells() As String)
            Rows.Add(New sRow(Cells))
        End Sub

        '''<summary>Write the current content of the cell buffer to a file.</summary>
        '''<param name="FileName">File to write to.</param>
        '''<returns>TRUE if write was ok, FALSE else.</returns>
        Public Function WriteAll(ByVal FileName As String) As Boolean

            Dim OutStream As New System.IO.StreamWriter(FileName)

            For Each Row As sRow In Rows
                OutStream.WriteLine(Join(Row.Cells, ";"))
            Next Row

            OutStream.Flush()
            OutStream.Close()

            Return True

        End Function


    End Class

End Namespace