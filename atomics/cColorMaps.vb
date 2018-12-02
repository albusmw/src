Option Explicit On
Option Strict On

Public Class cColorMaps

    Public Enum eMaps
        None
        Hot
        Jet
        Bone
    End Enum

    '''<summary>Invalid color in NONE schema.</summary>
    Private Shared ColorInvalid_None As Color = Color.Red
    '''<summary>Invalid color in HOT schema.</summary>
    Private Shared ColorInvalid_Hot As Color = Color.Blue
    '''<summary>Invalid color in JET schema.</summary>
    Private Shared ColorInvalid_Jet As Color = Color.Black
    '''<summary>Invalid color in BONE schema.</summary>
    Private Shared ColorInvalid_Bone As Color = Color.Red

    '''<summary>Calculate a color according to the grayscale schema.</summary>
    '''<param name="X">Value to calculate color for (0..255).</param>
    '''<returns>Color according to the value.</returns>
    Public Shared Function None(ByVal X As Double) As Drawing.Color
        If (X <= 255) And (X >= 0) Then
            Dim ColorAsByte As Byte = CByte(X)
            Return Drawing.Color.FromArgb(ColorAsByte, ColorAsByte, ColorAsByte)
        Else
            Return ColorInvalid_None
        End If
    End Function

    '''<summary>Calculate a color according to the hot schema (Black-red-yellow-white color map).</summary>
    '''<param name="X">Value to calculate color for (0..255).</param>
    '''<returns>Color according to the value.</returns>
    Public Shared Function Hot(ByVal X As Double) As Drawing.Color

        X /= 255

        Dim R As Byte, G As Byte, B As Byte

        If (X <= 1) And (X >= 0) Then
            If X >= 1 Then
                Return Drawing.Color.FromArgb(255, 255, 255)
            Else
                If X <= 0 Then
                    Return Drawing.Color.FromArgb(0, 0, 0)
                Else
                    R = CByte(255 * (CDbl(IIf(X < 0.4, X * 2.5, 1))))
                    G = CByte(255 * (CDbl(IIf(X >= 0.8, 1, IIf(X < 0.4, 0, (2.5 * X) - 1)))))
                    B = CByte(255 * (CDbl(IIf(X >= 0.8, (5 * X) - 4, 0))))
                    Return Drawing.Color.FromArgb(R, G, B)
                End If
            End If
        Else
            Return ColorInvalid_Hot
        End If

    End Function

    '''<summary>Calculate a color according to the jet schema (Variant of HSV).</summary>
    '''<param name="X">Value to calculate color for (0..255).</param>
    '''<returns>Color according to the value.</returns>
    Public Shared Function Jet(ByVal X As Double) As Drawing.Color

        Dim R As Byte, G As Byte, B As Byte

        X /= 255

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
            Case Else
                Return ColorInvalid_Jet
        End Select

        Return Drawing.Color.FromArgb(R, G, B)

    End Function

    '''<summary>Calculate a color according to the bone schema (Gray-scale with a tinge of blue color map).</summary>
    '''<param name="X">Value to calculate color for (0..255).</param>
    '''<returns>Color according to the value.</returns>
    Public Shared Function Bone(ByVal X As Double) As Drawing.Color

        X /= 255

        Dim R As Byte, G As Byte, B As Byte
        Dim C1 As Double = 29 / 24

        If (X <= 1) And (X >= 0) Then
            R = CByte(255 * CDbl(IIf(X < 0.75, X * 0.875, (1.375 * X) - 0.375)))
            G = CByte(255 * CDbl(IIf(X < 0.375, 0.875 * X, IIf(X >= 0.75, (0.875 * X) + 0.125, (C1 * X) - 0.125))))
            B = CByte(255 * CDbl(IIf(X < 0.375, C1 * X, (0.875 * X) + 0.125)))
            Return Drawing.Color.FromArgb(R, G, B)
        Else
            Return ColorInvalid_Bone
        End If

    End Function

End Class
