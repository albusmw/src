Option Explicit On
Option Strict On

''' <summary>Read and write image meta data.</summary>
''' <remarks>Needs reference to PresentationCore.dll and WindowsBase.</remarks>
Public Class cImageMetaData

    ''' <summary></summary>
    ''' <seealso cref="https://stackoverflow.com/questions/17590952/edit-iptc-metadata-on-jpg-vb-net/18195577"/>
    Public Shared Sub AddIPTCData(ByVal File As String)

        Dim Keywords As String = String.Empty
        Dim Description As String = String.Empty

        Dim stream As New IO.FileStream(File, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
        Dim decoder As New System.Windows.Media.Imaging.JpegBitmapDecoder(stream, Windows.Media.Imaging.BitmapCreateOptions.None, Windows.Media.Imaging.BitmapCacheOption.None)
        Dim metadata As Windows.Media.Imaging.BitmapMetadata = TryCast(decoder.Frames(0).Metadata, Windows.Media.Imaging.BitmapMetadata)
        If IsNothing(metadata) = False Then
            If IsNothing(metadata.Keywords) = False Then
                Keywords = metadata.Keywords.Aggregate(Function(old, val) Convert.ToString(old) & "; " & Convert.ToString(val))
            End If
        End If

        Description = CType(metadata.GetQuery("/app13/irb/8bimiptc/iptc/Caption"), String)                         'get description

        Dim bitmapFrame As Windows.Media.Imaging.BitmapFrame = decoder.Frames(0)
        metadata = CType(bitmapFrame.Metadata.Clone(), Windows.Media.Imaging.BitmapMetadata)
        Dim newkeywords As New List(Of String)(New String() {"test1", "test2"})
        If IsNothing(metadata.Keywords) = False Then
            newkeywords.AddRange(metadata.Keywords)                                                                 'this string adds old keywords
        End If

        metadata.Keywords = New ObjectModel.ReadOnlyCollection(Of String)(newkeywords)                              'replace keywords
        metadata.ApplicationName = "FITSWork - Lightroom - Self-made software"
        metadata.Author = New ObjectModel.ReadOnlyCollection(Of String)({"Martin Weiss as is"})
        metadata.CameraManufacturer = "QHY"
        metadata.CameraModel = "QHY600"
        metadata.Comment = "--- no comment ---"
        metadata.Copyright = "(c) 2019 Martin Weiss, www.sternwarte-holzkirchen.de"
        metadata.DateTaken = "12/21/2019"
        metadata.Subject = "NGC1234 - Iris Nebular"
        metadata.Title = "Test shot of NGC1234"

        metadata.SetQuery("/app13/irb/8bimiptc/iptc/Caption", "My test picture1.")                                  'set new description
        metadata.SetQuery("/app13/irb/8bimiptc/iptc/Caption", "My test picture1.")                                  'set new description

        Dim memstream As New IO.MemoryStream()                                                                      'create temp storage in memory
        Dim encoder As New System.Windows.Media.Imaging.JpegBitmapEncoder()
        encoder.Frames.Add(Windows.Media.Imaging.BitmapFrame.Create(bitmapFrame, bitmapFrame.Thumbnail, metadata, bitmapFrame.ColorContexts))
        encoder.Save(memstream) ' save in memory
        stream.Close()
        stream = New IO.FileStream(File, IO.FileMode.Open, IO.FileAccess.Write, IO.FileShare.ReadWrite)
        memstream.Seek(0, System.IO.SeekOrigin.Begin)                                                               'go to stream start
        Dim bytes(CInt(memstream.Length) - 1) As Byte
        memstream.Read(bytes, 0, CInt(memstream.Length))
        stream.Write(bytes, 0, bytes.Length)
        stream.Close()
        memstream.Close()

    End Sub

End Class