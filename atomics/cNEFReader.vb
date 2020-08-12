Option Explicit On
Option Strict On

'We use LibRaw for reading in Nikon NEF files
' -> https://www.libraw.org/docs/Samples-LibRaw.html

Public Class cNEFReader

    Public Sub Read(ByVal FileName As String)

        Dim Folder As String = "C:\Bin\LibRaw\bin"
        Dim EXE As String = "unprocessed_raw.exe"
        Dim Argument As String = "-T"

        Process.Start(System.IO.Path.Combine(Folder, EXE), Argument & " " & Chr(34) & FileName & Chr(34))

    End Sub

End Class