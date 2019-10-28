Option Explicit On
Option Strict On

Public Class SignalProcessing

    Public Shared Sub RegressPoly(ByRef X As Double(), ByRef Y As Double(), ByVal Order As Integer, ByRef Polynom As Double())

        Dim MomentMatrix As Double(,) = {}
        Dim RightSide As Double() = {}

        CalcFittingMatrix(X, Y, Order, MomentMatrix, RightSide)
        Polynom = MatrixSolve(MomentMatrix, RightSide)

    End Sub

    '''<summary>Calculate the polynomial fitting matrix required for the polynomial regression.</summary>
    '''<param name="X">Vector of X elements.</param>
    '''<param name="Y">Vector of Y elements.</param>
    '''<param name="Order">Requested order.</param>
    '''<param name="MomentMatrix">Matrix of moments.</param>
    '''<param name="RightSide">Right-side vector.</param>
    Public Shared Sub CalcFittingMatrix(ByRef X As Double(), ByRef Y As Double(), ByVal Order As Integer, ByRef MomentMatrix As Double(,), ByRef RightSide As Double())

        Dim MomentsVector As Double() = {}

        'Init return arguments
        ReDim MomentMatrix(Order, Order)
        ReDim RightSide(Order)

        'Calculate moments of the matrix
        GetPowerSums(X, MomentsVector, 2 * Order)

        'Fill in the moments in the matrix - upper left
        For Index As Integer = 0 To Order
            For IntIndex As Integer = 0 To Index
                MomentMatrix(IntIndex, Index - IntIndex) = MomentsVector(Index)
            Next IntIndex
        Next Index

        'Fill in the moments in the matrix - lower right
        For Index As Integer = Order + 1 To 2 * Order
            For IntIndex As Integer = Index - Order To Order
                MomentMatrix(IntIndex, Index - IntIndex) = MomentsVector(Index)
            Next IntIndex
        Next Index

        'Calculate right side
        For Index As Integer = 0 To Order
            RightSide(Index) = ArrayMixMoment(Y, X, 1, Index)
        Next Index

    End Sub

    Public Shared Sub GetPowerSums(ByRef Vector As Double(), ByRef Powers As Double(), ByVal MaxPower As Integer)

        Dim Index As Integer
        Dim Temp As Double()

        ReDim Powers(MaxPower)

        Powers(0) = UBound(Vector) - LBound(Vector) + 1     'calculate sum(x^0) which is only the vector length
        Temp = CreateCopy(Vector)                           'generate vector containing x^1

        For Index = 1 To Powers.GetUpperBound(0)
            Powers(Index) = Sum(Temp)                       'calculate the sum
            Mul_I(Temp, Vector)                             'multiply vector with x which makes x^n to x^(n+1)
        Next Index

    End Sub

    Public Shared Function ArrayMixMoment(ByRef Vector1 As Double(), ByRef Vector2 As Double(), ByVal Order1 As Integer, ByVal Order2 As Integer) As Double

        Dim Index As Long
        Dim Temp1 As Double() = {}
        Dim Temp2 As Double() = {}

        ReDim Temp1(Vector1.GetUpperBound(0))
        Init(Temp1, CDbl(1))
        If Order1 > 0 Then
            For Index = 1 To Order1
                Mul_I(Temp1, Vector1)
            Next Index
        End If

        ReDim Temp2(Vector2.GetUpperBound(0))
        Init(Temp2, CDbl(1))
        If Order2 > 0 Then
            For Index = 1 To Order2
                Mul_I(Temp2, Vector2)
            Next Index
        End If

        Return Sum(Mul(Temp1, Temp2))

    End Function

    '''<summary>Multiply 2 vectors.</summary>
    '''<param name="Vector1">1st vector to be multiplied .</param>
    '''<param name="Vector2">1st vector to be multiplied .</param>
    '''<returns>Multiplied vectors.</returns>
    Public Shared Function Mul(ByRef Vector1() As Double, ByVal Vector2() As Double) As Double()
        If IsNothing(Vector1) = True Then Return New Double() {}
        If IsNothing(Vector2) = True Then Return New Double() {}
        Dim RetVal As Double() : ReDim RetVal(Vector1.GetUpperBound(0))
        For Idx As Integer = 0 To RetVal.GetUpperBound(0)
            RetVal(Idx) = Vector1(Idx) * Vector2(Idx)
        Next Idx
        Return RetVal
    End Function

    Public Shared Function ArrayMixMoment(ByRef Vector1 As Double(), ByVal Order2 As Integer) As Double

        Dim RetVal As Double = 0

        For Idx As Integer = 0 To Vector1.GetUpperBound(0)
            RetVal += Vector1(Idx) * (Idx ^ Order2)
        Next Idx

        Return RetVal

    End Function

    '''<summary>Create a copy of the given vector.</summary>
    '''<param name="Vector">Vector to copy.</param>
    '''<returns>Copied vector.</returns>
    Public Shared Function CreateCopy(Of T)(ByRef Vector() As T) As T()
        If IsNothing(Vector) = True Then Return Nothing
        Dim RetVal(0 To Vector.GetUpperBound(0)) As T
        Array.Copy(Vector, 0, RetVal, 0, Vector.Length)
        Return RetVal
    End Function

    Public Shared Function Sum(ByRef Vector As Double()) As Double
        Dim RetVal As Double = 0
        For Idx As Integer = 0 To Vector.GetUpperBound(0)
            RetVal += Vector(Idx)
        Next Idx
        Return RetVal
    End Function

    '''<summary>Multiply 2 vectors in-place.</summary>
    '''<param name="InPlaceVector">Vector to be multiplied inplace.</param>
    '''<param name="VectorToMultiply">Vector that it multiplied to the given vector.</param>
    Public Shared Sub Mul_I(ByRef InPlaceVector() As Double, ByVal VectorToMultiply() As Double)
        If IsNothing(InPlaceVector) = True Then Exit Sub
        If IsNothing(VectorToMultiply) = True Then Exit Sub
        For Idx As Integer = 0 To InPlaceVector.GetUpperBound(0)
            InPlaceVector(Idx) *= VectorToMultiply(Idx)
        Next Idx
    End Sub

    '''<summary>Initialize a vector with a given value.</summary>
    '''<param name="InitValue">Initialization value.</param>
    '''<param name="Vector">Vector to copy.</param>
    Public Shared Sub Init(Of T)(ByRef Vector() As T, ByRef InitValue As T)
        If IsNothing(Vector) = False Then
            For Idx As Integer = 0 To Vector.GetUpperBound(0)
                Vector(Idx) = InitValue
            Next Idx
        End If
    End Sub

    Public Shared Function MatrixSolve(ByRef a As Double(,), ByRef B As Double()) As Double()

        REM Löst das Gleichungssystem A*x = b auf
        REM Strategie: Pivotisierung ohne Optimierungen

        Dim Index As Integer
        Dim IntIndex As Integer
        Dim HIndex As Integer
        Dim IntA As Double(,) = {}
        Dim IntB As Double() = {}
        Dim IntResult As Double() = {}
        Dim Upper As Integer
        Dim Teiler As Double

        Upper = UBound(a, 1)

        ReDim IntResult(Upper)

        ReDim IntA(a.GetUpperBound(0), a.GetUpperBound(1))
        Array.Copy(a, IntA, a.LongLength)
        IntB = CreateCopy(B)

        For Index = 0 To Upper                                                                  'pivotisieren

            Teiler = IntA(Index, Index)                                                        'Diagonal-Element
            For IntIndex = Index To Upper
                IntA(Index, IntIndex) = IntA(Index, IntIndex) / Teiler                         'Diagonal-Element = 1
            Next IntIndex
            IntB(Index) = IntB(Index) / Teiler

            For IntIndex = Index + 1 To Upper
                Teiler = IntA(IntIndex, Index)                                                 'Elemente unter aktuellem Diagonalelement
                If Teiler <> 0 Then
                    For HIndex = Index To Upper
                        IntA(IntIndex, HIndex) = IntA(IntIndex, HIndex) / Teiler               'unter Diagonale = 1
                    Next HIndex
                    IntB(IntIndex) = IntB(IntIndex) / Teiler
                End If
            Next IntIndex

            For IntIndex = Index + 1 To Upper                                                  'unter Diagonale = 0
                If IntA(IntIndex, Index) <> 0 Then
                    For HIndex = Index To Upper
                        IntA(IntIndex, HIndex) = IntA(IntIndex, HIndex) - IntA(Index, HIndex)
                    Next HIndex
                    IntB(IntIndex) = IntB(IntIndex) - IntB(Index)
                End If
            Next IntIndex

        Next Index

        REM AB HIER STEHEN AUF DER HAUPTDIAGONALE 1-EN UND IN DER UNTEREN LINKEN DREIECKSMATRIX 0-EN
        REM jetzt kann aufgelöst werden

        IntResult(Upper) = IntB(Upper)                                                          'Lösung letzte Zeile
        For Index = Upper - 1 To 0 Step -1                                                  'Zeilen up
            For IntIndex = Index + 1 To Upper
                IntResult(Index) = IntResult(Index) + (IntA(Index, IntIndex) * IntResult(IntIndex))
            Next IntIndex
            IntResult(Index) = IntB(Index) - IntResult(Index)
        Next Index

        MatrixSolve = IntResult

    End Function

End Class
