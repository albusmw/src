Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    Public Class Statistics

        '''<summary>Calculate max and min value of the given matrix.</summary>
        '''<param name="Matrix">Matrix to calculate statistics for.</param>
        '''<param name="Min">Detected MIN value.</param>
        '''<param name="Max">Detected MAX value.</param>
        Public Shared Sub MaxMin(ByRef Matrix(,) As Double, ByRef Min As Double, ByRef Max As Double)
            Max = Double.MinValue : Min = Double.MaxValue
            For Idx1 As Integer = 0 To Matrix.GetUpperBound(0)
                For Idx2 As Integer = 0 To Matrix.GetUpperBound(1)
                    If Matrix(Idx1, Idx2) > Max Then Max = Matrix(Idx1, Idx2)
                    If Matrix(Idx1, Idx2) < Min Then Min = Matrix(Idx1, Idx2)
                Next Idx2
            Next Idx1
        End Sub

        '''<summary>Find the peak position.</summary>
        '''<param name="Matrix">Matrix to calculate statistics for.</param>
        Public Shared Sub FindPeak(ByRef Matrix(,) As Double, ByRef X As Integer, ByRef Y As Integer)
            X = -1 : Y = -1
            Dim Max As Double = Double.MinValue
            For Idx1 As Integer = 0 To Matrix.GetUpperBound(0)
                For Idx2 As Integer = 0 To Matrix.GetUpperBound(1)
                    If Matrix(Idx1, Idx2) > Max Then
                        Max = Matrix(Idx1, Idx2) : X = Idx1 : Y = Idx2
                    End If
                Next Idx2
            Next Idx1
        End Sub

        '''<summary>Calculate mean value of the given matrix.</summary>
        '''<param name="Matrix">Matrix to calculate statistics for.</param>
        Public Shared Function Mean(ByRef Matrix(,) As Double) As Double
            Dim RetVal As Double = 0
            For Idx1 As Integer = 0 To Matrix.GetUpperBound(0)
                For Idx2 As Integer = 0 To Matrix.GetUpperBound(1)
                    RetVal += Matrix(Idx1, Idx2)
                Next Idx2
            Next Idx1
            Return RetVal / (Matrix.LongLength)
        End Function

    End Class

End Namespace