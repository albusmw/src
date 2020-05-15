Option Explicit On
Option Strict On

Public Class PictureBoxEx : Inherits System.Windows.Forms.PictureBox

    Public Property InterpolationMode As Drawing.Drawing2D.InterpolationMode

    Protected Overrides Sub OnPaint(ByVal paintEventArgs As Windows.Forms.PaintEventArgs)
        paintEventArgs.Graphics.InterpolationMode = InterpolationMode
        MyBase.OnPaint(paintEventArgs)
    End Sub

    Public Shadows Event MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs)

    Private LastMouseLocation As Drawing.Point = Nothing

    Public Sub MouseMove_inner(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        LastMouseLocation = e.Location
        RaiseEvent MouseMove(sender, e)
    End Sub

    '''<summary>Translate the coordinates of the zoom mode image to "real" image coordinates</summary>
    '''<param name="Coordinates">Coordinates to translate.</param>
    '''<returns>Coordinates within the image.</returns>
    Public Function ScreenCoordinatesToImageCoordinates() As Drawing.PointF
        Return ScreenCoordinatesToImageCoordinates(LastMouseLocation)
    End Function

    '''<summary>Translate the coordinates of the zoom mode image to "real" image coordinates</summary>
    '''<param name="Coordinates">Coordinates to translate.</param>
    '''<returns>Coordinates within the image.</returns>
    Public Function ScreenCoordinatesToImageCoordinates(ByVal Coordinates As Drawing.Point) As Drawing.PointF

        'Test to make sure our image is not null
        If Me.Image Is Nothing Then Return Coordinates

        'Make sure our control width and height are not 0 and our image width and height are not 0
        If Me.Width = 0 OrElse Me.Height = 0 OrElse Me.Image.Width = 0 OrElse Me.Image.Height = 0 Then Return Coordinates

        ' This is the one that gets a little tricky. Essentially, need to check 
        ' the aspect ratio of the image to the aspect ratio of the control
        ' to determine how it is being rendered
        Dim imageAspect As Double = Me.Image.Width / Me.Image.Height
        Dim controlAspect As Double = Me.Width / Me.Height
        Dim newX As Double = Double.NaN
        Dim newY As Double = Double.NaN

        If imageAspect > controlAspect Then

            'Image has "black bars on top and bottom"
            newX = CSng(Me.Image.Width * (Coordinates.X / Me.Width))                'Width if fully used -> direct convert

            Dim DispImageHeight As Double = Me.Width / imageAspect                  'image height [display pixel]
            Dim BlackBarHeight As Double = (Me.Height - DispImageHeight) / 2        'height of black bar on top [display pixel]

            newY = (Coordinates.Y - BlackBarHeight)                                 'Y coordinate starts at 0 in image top
            newY = newY * (Me.Image.Height / DispImageHeight)                       'scale from height [display pixel] to height [image pixel]

        Else

            'Image has "black bars on left and right"
            newY = CSng(Me.Image.Height * (Coordinates.Y / Me.Height))              'Height if fully used -> direct convert

            Dim DispImageWidth As Double = Me.Height * imageAspect                  'image width [display pixel]
            Dim BlackBarWidth As Double = (Me.Width - DispImageWidth) / 2           'width of black bar on left [display pixel]

            newX = (Coordinates.X - BlackBarWidth)                                  'X coordinate starts at 0 in image left
            newX = newX * (Me.Image.Width / DispImageWidth)                         'scale from width [display pixel] to width [image pixel]

        End If

        Return New Drawing.PointF(CSng(newX), CSng(newY))

    End Function

    '''<summary>Get the coordinates to zoom in.</summary>
    '''<param name="Center">Center point in image pixel coordinates.</param>
    '''<remarks>Y_top has a lower value compared to Y_bottom (we count top-down).</remarks>
    Public Shared Function CenterSizeToXY(ByVal Center As Drawing.PointF, ByVal Size As Integer, ByRef X_left As Integer, ByRef X_right As Integer, ByRef Y_top As Integer, ByRef Y_bottom As Integer) As Drawing.Point

        Dim Mod2 As Integer = CInt(IIf(Size Mod 2 = 1, 1, 0))
        X_left = CInt(Center.X) - (CInt((Size - Mod2) / 2))
        X_right = CInt(Center.X) + (CInt((Size - Mod2) / 2))
        Y_top = CInt(Center.Y) - (CInt((Size - Mod2) / 2))
        Y_bottom = CInt(Center.Y) + (CInt((Size - Mod2) / 2))

        Return New Drawing.Point(CInt(Center.X), CInt(Center.Y))

    End Function

End Class