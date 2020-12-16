Option Explicit On
Option Strict On

'We use LibRaw for reading in Nikon NEF files
' -> https://www.libraw.org/docs/Samples-LibRaw.html

Public Class cNEFReader

    Public Function Read(ByVal NEFFileLocation As String) As String

        Dim Folder As String = "C:\Bin\LibRaw\bin"
        Dim EXE As String = "unprocessed_raw.exe"
        Dim Argument As String = "-T"

        'Copy file to temporary location
        If System.IO.File.Exists(NEFFileLocation) = False Then Return "File <" & NEFFileLocation & "> not found"
        Dim FileNameOnly As String = System.IO.Path.GetFileName(NEFFileLocation)
        Dim TempFile As String = System.IO.Path.Combine(System.IO.Path.GetTempPath, FileNameOnly)
        If System.IO.File.Exists(TempFile) Then System.IO.File.Delete(TempFile)
        System.IO.File.Copy(NEFFileLocation, TempFile)

        'Run LibRaw
        Dim LibRawStartInfo As New ProcessStartInfo
        With LibRawStartInfo
            .FileName = System.IO.Path.Combine(Folder, EXE)
            .Arguments = Argument & " " & Chr(34) & TempFile & Chr(34)
            .UseShellExecute = False
            .RedirectStandardOutput = True
            .RedirectStandardError = True
        End With

        Dim LibRaw As Process = Process.Start(LibRawStartInfo)
        LibRaw.WaitForExit()
        Dim OutInfo As String = LibRaw.StandardOutput.ReadToEnd

    End Function

End Class