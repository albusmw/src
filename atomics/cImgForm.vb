Option Explicit On
Option Strict On

'''<summary>Display a simple picturebox window (form and graph on it).</summary>
Public Class cImgForm

    '''<summary>The form that shall be displayed.</summary>
    Public Hoster As System.Windows.Forms.Form = Nothing
    '''<summary>The ZED graph control inside the form.</summary>
    Public Image As PictureBoxEx = Nothing

    '''<summary>Prepare.</summary>
    Public Sub New()
        If IsNothing(Hoster) = True Then Hoster = New System.Windows.Forms.Form
        If IsNothing(Image) = True Then
            Image = New PictureBoxEx
            Hoster.Controls.Add(Image)
            With Image
                .Dock = Windows.Forms.DockStyle.Fill
                .InterpolationMode = Drawing.Drawing2D.InterpolationMode.NearestNeighbor
                .SizeMode = Windows.Forms.PictureBoxSizeMode.Zoom
                .BackColor = Drawing.Color.Black
            End With
        End If
    End Sub

    Public Function Show(ByVal NewTitle As String) As System.Windows.Forms.PictureBox
        Hoster.Text = NewTitle
        Hoster.Show()
        Return Image
    End Function

    '''<summary>Update the content of the focus window.</summary>
    '''<param name="Form">Focus window.</param>
    '''<param name="Data">Data to display.</param>
    '''<param name="MaxData">Maximum in the data in order to normalize correct.</param>
    Public Sub ShowData(ByRef Data(,) As UInt16, ByVal MinData As Long, ByVal MaxData As Long)
        Dim OutputImage As New cLockBitmap(Data.GetUpperBound(0), Data.GetUpperBound(1))
        If MaxData = 0 Then MaxData = 1
        If MaxData = MinData Then Exit Sub
        OutputImage.LockBits()
        Dim Stride As Integer = OutputImage.BitmapData.Stride
        Dim BytePerPixel As Integer = OutputImage.ColorBytesPerPixel
        Dim YOffset As Integer = 0
        For Y As Integer = 0 To OutputImage.Height - 1
            Dim BaseOffset As Integer = YOffset
            For X As Integer = 0 To OutputImage.Width - 1
                Dim DispVal As Integer = CInt((Data(X, Y) - MinData) * (255 / (MaxData - MinData)))
                Dim Coloring As Drawing.Color = cColorMaps.Jet(DispVal)
                OutputImage.Pixels(BaseOffset) = Coloring.R
                OutputImage.Pixels(BaseOffset + 1) = Coloring.G
                OutputImage.Pixels(BaseOffset + 2) = Coloring.B
                BaseOffset += BytePerPixel
            Next X
            YOffset += Stride
        Next Y
        OutputImage.UnlockBits()
        Image.Image = OutputImage.BitmapToProcess
    End Sub

    '''<summary>Update the content of the focus window.</summary>
    '''<param name="Form">Focus window.</param>
    '''<param name="Data">Data to display.</param>
    '''<param name="MaxData">Maximum in the data in order to normalize correct.</param>
    Public Sub ShowData(ByRef Data(,) As UInt32, ByVal MinData As Long, ByVal MaxData As Long)
        Dim OutputImage As New cLockBitmap(Data.GetUpperBound(0), Data.GetUpperBound(1))
        If MaxData = 0 Then MaxData = 1
        OutputImage.LockBits()
        Dim Stride As Integer = OutputImage.BitmapData.Stride
        Dim BytePerPixel As Integer = OutputImage.ColorBytesPerPixel
        Dim YOffset As Integer = 0
        For Y As Integer = 0 To OutputImage.Height - 1
            Dim BaseOffset As Integer = YOffset
            For X As Integer = 0 To OutputImage.Width - 1
                Dim DispVal As Integer = CInt((Data(X, Y) - MinData) * (255 / (MaxData - MinData)))
                Dim Coloring As Drawing.Color = cColorMaps.Jet(DispVal)
                OutputImage.Pixels(BaseOffset) = Coloring.R
                OutputImage.Pixels(BaseOffset + 1) = Coloring.G
                OutputImage.Pixels(BaseOffset + 2) = Coloring.B
                BaseOffset += BytePerPixel
            Next X
            YOffset += Stride
        Next Y
        OutputImage.UnlockBits()
        Image.Image = OutputImage.BitmapToProcess
    End Sub

End Class