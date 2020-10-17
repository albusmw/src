Option Explicit On
Option Strict On

Public Class GZIP

  Public Shared Function Decompress(ByVal FileName As String) As Byte()

    Dim MemStream As New IO.MemoryStream(System.IO.File.ReadAllBytes(FileName))
    Dim stream As New IO.Compression.GZipStream(MemStream, IO.Compression.CompressionMode.Decompress)

    Const size As Integer = 4096
    Dim buffer = New Byte(size - 1) {}
    Dim Uncompressed As New IO.MemoryStream()
    Dim count As Integer
    Do
      count = stream.Read(buffer, 0, size)
      If count > 0 Then
        Uncompressed.Write(buffer, 0, count)
      End If
    Loop While count > 0
    Return Uncompressed.ToArray()

  End Function

  Public Shared Function DecompressTo(ByVal CompressedFile As String, ByVal DecompressedFile As String) As Boolean
        Try
            System.IO.File.WriteAllBytes(DecompressedFile, Decompress(CompressedFile))
            Return True
        Catch ex As Exception
            Return False
        End Try
  End Function

End Class
