Option Explicit On
Option Strict On

Namespace Colors

    '''<summary>This class can be used to generate a sequence of colors, e.g. for a plot.</summary>
    Public Class cColorList

        Private ColorList As List(Of System.Drawing.Color)
        Private ColorPtr As Integer = -1

        Public Sub New()

            ColorPtr = -1
            ColorList = New List(Of System.Drawing.Color)

            ColorList.Add(Drawing.Color.Red)
            ColorList.Add(Drawing.Color.Green)
            ColorList.Add(Drawing.Color.Blue)
            ColorList.Add(Drawing.Color.Cyan)
            ColorList.Add(Drawing.Color.Magenta)
            ColorList.Add(Drawing.Color.Yellow)

            ColorList.Add(Drawing.Color.DarkRed)
            ColorList.Add(Drawing.Color.DarkGreen)
            ColorList.Add(Drawing.Color.DarkBlue)
            ColorList.Add(Drawing.Color.DarkCyan)
            ColorList.Add(Drawing.Color.DarkMagenta)
            ColorList.Add(Drawing.Color.LightYellow)

        End Sub

        Public Function GetNextColor() As System.Drawing.Color
            ColorPtr += 1
            If ColorPtr > ColorList.Count - 1 Then ColorPtr = 0
            Return ColorList(ColorPtr)
        End Function

    End Class

    Public Class ColorConversion

        Public Shared Function ToBrush(ByVal Color As System.Drawing.Color) As System.Drawing.Brush
            Return New System.Drawing.SolidBrush(Color)
        End Function

        '''<summary>Convert a color from HSB to RGB system.</summary>
        '''<param name="Alpha">Alpha blending value {0..255} - not modified, just passed to the produced color.</param>
        '''<param name="Hue">Coloring {0..360}.</param>
        '''<param name="Saturation">Saturation {0..1}.</param>
        '''<param name="Brightness">Brightness {0..1}.</param>
        '''<returns>RGB color.</returns>
        Public Shared Function ColorFromAHSB(ByVal Alpha As Integer, ByVal Hue As Single, ByVal Saturation As Single, ByVal Brightness As Single) As Drawing.Color

            '"Out of range" conversions
            If 0 > Alpha Or 255 < Alpha Then Return Drawing.Color.Black
            If 0 > Hue Or 360 < Hue Then Return Drawing.Color.Black
            If 0 > Saturation Or 1 < Saturation Then Return Drawing.Color.Black
            If 0 > Brightness Or 1 < Brightness Then Return Drawing.Color.Black

            'No color saturation results in gray scale.
            If Saturation = 0 Then
                Dim GrayVal As Integer = Convert.ToInt32(Brightness * 255)
                Return Drawing.Color.FromArgb(Alpha, GrayVal, GrayVal, GrayVal)
            End If

            Dim fMax, fMid, fMin As Single
            Dim iSextant, iMax, iMid, iMin As Integer

            If Brightness > 0.5 Then
                fMax = Brightness - (Brightness * Saturation) + Saturation
                fMin = Brightness + (Brightness * Saturation) - Saturation
            Else
                fMax = Brightness + (Brightness * Saturation)
                fMin = Brightness - (Brightness * Saturation)
            End If

            iSextant = CInt(System.Math.Floor(Hue / 60))
            If 300 <= Hue Then
                Hue -= 360.0F
            End If

            Hue /= 60
            Hue = CSng(Hue - 2.0F * System.Math.Floor(((iSextant + 1.0F) Mod 6.0F) / 2.0F))
            If iSextant Mod 2 = 0 Then
                fMid = Hue * (fMax - fMin) + fMin
            Else
                fMid = fMin - Hue * (fMax - fMin)
            End If

            iMax = Convert.ToInt32(fMax * 255)
            iMid = Convert.ToInt32(fMid * 255)
            iMin = Convert.ToInt32(fMin * 255)

            Select Case iSextant
                Case 1
                    Return Drawing.Color.FromArgb(Alpha, iMid, iMax, iMin)
                Case 2
                    Return Drawing.Color.FromArgb(Alpha, iMin, iMax, iMid)
                Case 3
                    Return Drawing.Color.FromArgb(Alpha, iMin, iMid, iMax)
                Case 4
                    Return Drawing.Color.FromArgb(Alpha, iMid, iMin, iMax)
                Case 5
                    Return Drawing.Color.FromArgb(Alpha, iMax, iMin, iMid)
                Case Else
                    Return Drawing.Color.FromArgb(Alpha, iMax, iMid, iMin)
            End Select

        End Function

    End Class

    '''<summary>This class supports the conversion between values / 2D matrix elements to different color schemas</summary>
    '''<remarks>Common example: X ray pictures code the transmission intensity with gray scale.</remarks>
    Public Class ColorSchema

        '''<summary>Calculate a color according to the grayscale schema (the lighter the whiter).</summary>
        '''<param name="X">Value to calculate color for (0..1).</param>
        '''<returns>Color according to the value.</returns>
        Public Shared Function Gray(ByVal X As Double) As Drawing.Color
            Dim GrayValue As Byte = CByte(X * 255)
            Return Drawing.Color.FromArgb(GrayValue, GrayValue, GrayValue)
        End Function

        '''<summary>Calculate a color according to the bone schema (Gray-scale with a tinge of blue color map).</summary>
        '''<param name="X">Value to calculate color for (0..1).</param>
        '''<returns>Color according to the value.</returns>
        Public Shared Function Bone(ByVal X As Double) As Drawing.Color

            Dim R As Byte, G As Byte, B As Byte
            Dim C1 As Double = 29 / 24

            R = CByte(255 * CDbl(IIf(X < 0.75, X * 0.875, (1.375 * X) - 0.375)))
            G = CByte(255 * CDbl(IIf(X < 0.375, 0.875 * X, IIf(X >= 0.75, (0.875 * X) + 0.125, (C1 * X) - 0.125))))
            B = CByte(255 * CDbl(IIf(X < 0.375, C1 * X, (0.875 * X) + 0.125)))
            Return Drawing.Color.FromArgb(R, G, B)

        End Function

        '''<summary>Calculate a color according to the copper schema (Linear copper-tone color map).</summary>
        '''<param name="X">Value to calculate color for (0..1).</param>
        '''<returns>Color according to the value.</returns>
        Public Shared Function Copper(ByVal X As Double) As Drawing.Color

            Dim R As Byte, G As Byte, B As Byte

            If X = 1 Then
                Return Drawing.Color.FromArgb(0, 0, 0)
            Else
                R = CByte(255 * (CDbl(IIf(X < 0.8, X * 1.25, 1))))
                G = CByte(255 * (0.8 * X))
                B = CByte(255 * (0.5 * X))
                Return Drawing.Color.FromArgb(R, G, B)
            End If

        End Function

        ''' <summary>
        ''' Calculate a color according to the hot schema (Black-red-yellow-white color map).
        ''' </summary>
        ''' <param name="X">Value to calculate color for (0..1).</param>
        ''' <returns>Color according to the value.</returns>
        Public Shared Function Hot(ByVal X As Double) As Drawing.Color

            Dim R As Byte, G As Byte, B As Byte

            If X = 1 Then
                Hot = Drawing.Color.FromArgb(0, 0, 0)
            Else
                R = CByte(255 * (CDbl(IIf(X < 0.4, X * 2.5, 1))))
                G = CByte(255 * (CDbl(IIf(X >= 0.8, 1, IIf(X < 0.4, 0, (2.5 * X) - 1)))))
                B = CByte(255 * (CDbl(IIf(X >= 0.8, (5 * X) - 4, 0))))
                Return Drawing.Color.FromArgb(R, G, B)
            End If

        End Function

        '''<summary>Calculate a color according to the jet schema (Variant of HSV).</summary>
        '''<param name="X">Value to calculate color for (0..1).</param>
        '''<returns>Color according to the value.</returns>
        Public Shared Function Jet(ByVal X As Double) As Drawing.Color

            Dim R As Byte, G As Byte, B As Byte

            Select Case X
                Case 0 To 0.125
                    R = 0
                    G = 0
                    B = CByte(255 * (0.5 + (4 * X)))
                Case 0.125 To 0.375
                    R = 0
                    G = CByte(255 * ((4 * X) - 0.5))
                    B = 255
                Case 0.375 To 0.625
                    R = CByte(255 * ((4 * X) - 1.5))
                    G = 255
                    B = CByte(255 * (2.5 - (4 * X)))
                Case 0.625 To 0.875
                    R = 255
                    G = CByte(255 * (3.5 - (4 * X)))
                    B = 0
                Case 0.875 To 1
                    R = CByte(255 * (4.5 - (4 * X)))
                    G = 0
                    B = 0
            End Select

            Return Drawing.Color.FromArgb(R, G, B)

        End Function

    End Class

End Namespace