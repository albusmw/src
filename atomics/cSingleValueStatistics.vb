Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Implements statistics for double precision values.</summary>
    '''<todo>Try to make this generic ...</todo>
    Public Class cSingleValueStatistics

        '''<summary>Type of values to apply statistics to.</summary>
        Public Enum eValueType
            '''<summary>Linear.</summary>
            Linear = 0
            '''<summary>Values are logarithmic, factor 10.</summary>
            LogBase10 = 1
            '''<summary>Values are logarithmic, factor 20.</summary>
            LogBase20 = 2
        End Enum

        Public Enum eAspects
            Maximum
            Minimum
            AbsMax
            Mean
            RMS
            Sigma
            ConfidenceMax
            ConfidenceMin
        End Enum

        '============================================================
        'PRIVATES
        '============================================================

        '''<summary>Internal values used for aggregation.</summary>
        Private Structure sIntProps
            '''<summary>Averaging type (lin, log10 or log20).</summary>
            Public ValueType As eValueType
            '''<summary>Number of values aggregated.</summary>
            Public ValueCount As Long
            '''<summary>Values that did NOT contribute to the statistics.</summary>
            Public InvalidValueCount As Long
            '''<summary>Maximum found.</summary>
            Public Maximum As Double
            '''<summary>Minimum found.</summary>
            Public Minimum As Double
            '''<summary>Linar sum of all value.</summary>
            Public Sum As Double
            '''<summary>Square sum of all value.</summary>
            Public SquareSum As Double
        End Structure
        Private IntProps As sIntProps

        '''<summary>Raw value storage (all input values go here if StoreValues is selected).</summary>
        Private RawValueStorage As New List(Of Double)
        '''<summary>Sorted values - calculated "on demand" if percentile is queried.</summary>
        Private Sorted As Double() = Nothing

        '''<summary>Name of the statistic value.</summary>
        Public Property Name() As String
            Get
                Return MyName
            End Get
            Set(ByVal value As String)
                MyName = value
            End Set
        End Property
        Private MyName As String = String.Empty

        '============================================================
        'PROPERTIES
        '============================================================

        '''<summary>Store all data that are pushed to the class.</summary>
        Public Property StoreRawValues() As Boolean
            Get
                Return MyStoreRawValues
            End Get
            Set(ByVal value As Boolean)
                MyStoreRawValues = value
            End Set
        End Property
        Private MyStoreRawValues As Boolean = False

        '''<summary>Multiplier of confidence values.</summary>
        Public Property ConfidenceMultiplier() As Double
            Get
                Return MyConfidenceMultiplier
            End Get
            Set(ByVal value As Double)
                MyConfidenceMultiplier = value
            End Set
        End Property
        Private MyConfidenceMultiplier As Double = 2

        Public ReadOnly Property ValueType() As eValueType
            Get
                Return IntProps.ValueType
            End Get
        End Property

        '''<summary>Number of values used to calculate statistics.</summary>
        Public ReadOnly Property ValueCount() As Long
            Get
                Return IntProps.ValueCount
            End Get
        End Property

        '''<summary>Number of invalid values tried to add (and are not added).</summary>
        Public ReadOnly Property InvalidValueCount() As Long
            Get
                Return IntProps.InvalidValueCount
            End Get
        End Property

        '================================================================================
        'Statistic results
        '================================================================================

        '''<summary>Maximum over all added values.</summary>
        Public ReadOnly Property Maximum() As Double
            Get
                If IntProps.ValueCount > 0 Then
                    Return IntProps.Maximum
                Else
                    Return Double.NaN
                End If
            End Get
        End Property

        '''<summary>Minimum over all added values.</summary>
        Public ReadOnly Property Minimum() As Double
            Get
                If IntProps.ValueCount > 0 Then
                    Return IntProps.Minimum
                Else
                    Return Double.NaN
                End If
            End Get
        End Property

        '''<summary>Maximum - Min over all added values.</summary>
        Public ReadOnly Property MaxMin() As Double
            Get
                If IntProps.ValueCount > 0 Then
                    Return IntProps.Maximum - IntProps.Minimum
                Else
                    Return Double.NaN
                End If
            End Get
        End Property

        '''<summary>Maximum absolute value with sign (AbsMax(-3,1)=-3).</summary>
        Public ReadOnly Property AbsMax() As Double
            Get
                If IntProps.ValueCount > 0 Then
                    If System.Math.Abs(IntProps.Minimum) > System.Math.Abs(IntProps.Maximum) Then
                        Return IntProps.Minimum
                    Else
                        Return IntProps.Maximum
                    End If
                Else
                    Return Double.NaN
                End If
            End Get
        End Property

        '''<summary>Mean value (calculation depends on value type (linear, log).</summary>
        Public ReadOnly Property Mean() As Double
            Get
                Dim LinMean As Double = IntProps.Sum / IntProps.ValueCount
                Select Case IntProps.ValueType
                    Case eValueType.Linear
                        Return LinMean
                    Case eValueType.LogBase10
                        Return 10 * System.Math.Log10(LinMean)
                    Case eValueType.LogBase20
                        Return 20 * System.Math.Log10(LinMean)
                    Case Else
                        Throw New Exception("Requested mean value type is not defined.")
                End Select
            End Get
        End Property

        '''<summary>RMS value (calculation depends on value type (linear, log).</summary>
        Public ReadOnly Property RMS() As Double
            Get
                Dim RMSInt As Double = System.Math.Sqrt(IntProps.SquareSum / IntProps.ValueCount)
                Select Case IntProps.ValueType
                    Case eValueType.Linear
                        Return RMSInt
                    Case eValueType.LogBase10
                        Return 10 * System.Math.Log10(RMSInt)
                    Case eValueType.LogBase20
                        Return 20 * System.Math.Log10(RMSInt)
                    Case Else
                        Throw New Exception("Requested RMS value type is not defined.")
                End Select
            End Get
        End Property

        '''<summary>Standard deviation (calculation depends on value type (linear, log).</summary>
        '''<todo>Check if the calculation of the logarithmic values are correct.</todo>
        Public ReadOnly Property Sigma() As Double
            Get
                Dim LinSigma As Double = System.Math.Sqrt(LinVariance)
                Dim LinMean As Double = IntProps.Sum / IntProps.ValueCount
                Select Case IntProps.ValueType
                    Case eValueType.Linear
                        Return LinSigma
                    Case eValueType.LogBase10
                        Return ((10 * System.Math.Log10(LinMean + LinSigma)) - (10 * System.Math.Log10(LinMean - LinSigma))) / 2
                    Case eValueType.LogBase20
                        Return ((20 * System.Math.Log10(LinMean + LinSigma)) - (20 * System.Math.Log10(LinMean - LinSigma))) / 2
                    Case Else
                        Throw New Exception("Requested sigma value type is not defined.")
                End Select
            End Get
        End Property

        '''<summary>Maximum of confidence interval (calculation depends on value type (linear, log).</summary>
        '''<param name="SpecialConfidenceMultiplier">Multiplier (1 = 65 %, 2 = 95 %, 3 = 99 %).</param>
        Public ReadOnly Property ConfidenceMax(ByVal SpecialConfidenceMultiplier As Double) As Double
            Get
                Return Mean + (Sigma * SpecialConfidenceMultiplier)
            End Get
        End Property

        '''<summary>Maximum of 95 % confidence interval (calculation depends on value type (linear, log)).</summary>
        Public ReadOnly Property ConfidenceMax() As Double
            Get
                Return ConfidenceMax(ConfidenceMultiplier)
            End Get
        End Property

        '''<summary>Minimum of confidence interval (calculation depends on value type (linear, log).</summary>
        '''<param name="SpecialConfidenceMultiplier">Multiplier (1 = 65 %, 2 = 95 %, 3 = 99 %).</param>
        Public ReadOnly Property ConfidenceMin(ByVal SpecialConfidenceMultiplier As Double) As Double
            Get
                Return Mean - (Sigma * SpecialConfidenceMultiplier)
            End Get
        End Property

        '''<summary>Minimum of 95 % confidence interval (calculation depends on value type (linear, log).</summary>
        Public ReadOnly Property ConfidenceMin() As Double
            Get
                Return ConfidenceMin(ConfidenceMultiplier)
            End Get
        End Property

        '''<summary>Returns the value for which 95 percent of the measured values are below this value.</summary>
        '''<returns>Value for which 95 percent of the measured values are below this value.</returns>
        Public ReadOnly Property Perc95() As Double
            Get
                Return Percentile(95)
            End Get
        End Property

        '''<summary>Returns the value for which 5 % of the measured values are below this value.</summary>
        '''<returns>Value for which 5 % of the measured values are below this value.</returns>
        Public ReadOnly Property Perc5() As Double
            Get
                Return Percentile(5)
            End Get
        End Property

        '''<summary>Calculate the percentile value.</summary>
        '''<param name="PercentileValue">Percentile (scale: 0 .. 100 %).</param>
        '''<returns>Value for which <paramref name="PercentileValue"/> % of the measured values are below this value.</returns>
        Public ReadOnly Property Percentile(ByVal PercentileValue As Double) As Double
            Get
                If RawValueStorage.Count = 0 Then Return Double.NaN 'if raw values are not stored -> return NaN
                If IsNothing(Sorted) Then                                   'if the sorted values are not available
                    Sorted = RawValues                                   'create a copy of the raw data
                    Array.Sort(Sorted)                                      'sort the data
                End If
                'Calculate the index to access and return the corresponding vector
                Dim IndexToAccess As Integer = CInt(Sorted.GetUpperBound(0) * (PercentileValue / 100))
                If IndexToAccess < 0 Then Return Sorted(0)
                If IndexToAccess > Sorted.GetUpperBound(0) Then Return Sorted(Sorted.GetUpperBound(0))
                Return Sorted(IndexToAccess)
            End Get
        End Property

        '''<summary>A vector of all raw values which had been stored (if StoreRawValues was selected).</summary>
        '''<returns>Vector of all raw values.</returns>
        Public ReadOnly Property RawValues() As Double()
            Get
                Return RawValueStorage.ToArray
            End Get
        End Property

        '''<summary>Add a complete vector to the statistic aggregator.</summary>
        '''<param name="Values">Value array to add.</param>
        Public Sub AddValueRange(ByRef Values() As Double)
            Sorted = Nothing
            For Idx As Integer = 0 To Values.GetUpperBound(0)
                AddValue(Values(Idx))
            Next Idx
        End Sub

        '''<summary>Add a complete vector to the statistic aggregator.</summary>
        '''<param name="Values">Value array to add.</param>
        Public Sub AddValueRange(ByRef Values() As Single)
            Sorted = Nothing
            For Idx As Integer = 0 To Values.GetUpperBound(0)
                AddValue(Values(Idx))
            Next Idx
        End Sub

        '''<summary>Add 1 new value to the statistics without NAN, INF, ... check.</summary>
        '''<param name="Value">Value to be added.</param>
        Public Sub AddValueUnsave(ByVal Value As Double)
            Sorted = Nothing
            If StoreRawValues = True Then RawValueStorage.Add(Value)
            Dim LinValue As Double
            IntProps.ValueCount += 1
            If Value > IntProps.Maximum Then IntProps.Maximum = Value Else If Value < IntProps.Minimum Then IntProps.Minimum = Value
            Select Case IntProps.ValueType
                Case eValueType.Linear
                    LinValue = Value
                Case eValueType.LogBase10
                    LinValue = 10 ^ (Value / 10)
                Case eValueType.LogBase20
                    LinValue = 10 ^ (Value / 20)
            End Select
            IntProps.Sum += LinValue
            IntProps.SquareSum += LinValue * LinValue
        End Sub

        '''<summary>Add 1 new value to the statistics.</summary>
        '''<param name="Value">Value to be added.</param>
        Public Sub AddValue(ByVal Value As Double)
            Sorted = Nothing
            If StoreRawValues = True Then RawValueStorage.Add(Value)
            If Double.IsNaN(Value) = False And Double.IsPositiveInfinity(Value) = False And Double.IsNegativeInfinity(Value) = False Then
                Dim LinValue As Double
                IntProps.ValueCount += 1
                If Value > IntProps.Maximum Then IntProps.Maximum = Value
                If Value < IntProps.Minimum Then IntProps.Minimum = Value
                Select Case IntProps.ValueType
                    Case eValueType.Linear
                        LinValue = Value
                    Case eValueType.LogBase10
                        LinValue = 10 ^ (Value / 10)
                    Case eValueType.LogBase20
                        LinValue = 10 ^ (Value / 20)
                End Select
                IntProps.Sum += LinValue
                IntProps.SquareSum += LinValue * LinValue
            Else
                IntProps.InvalidValueCount += 1
            End If
        End Sub

        '''<summary>Initialize the statistics calculator with the given type of elements.</summary>
        '''<param name="ValueType">Type of elements (linear, log, ...).</param>
        Public Sub New(ByVal Name As String, ByVal ValueType As eValueType)
            Me.New(ValueType, False)
            Me.Name = Name
        End Sub

        '''<summary>Initialize the statistics calculator with the given type of elements.</summary>
        '''<param name="ValueType">Type of elements (linear, log, ...).</param>
        Public Sub New(ByVal Name As String, ByVal ValueType As eValueType, ByVal StoreRawValues As Boolean)
            Me.New(ValueType, StoreRawValues)
            Me.Name = Name
        End Sub

        '''<summary>Initialize the statistics calculator with the given type of elements.</summary>
        '''<param name="ValueType">Type of elements (linear, log, ...).</param>
        Public Sub New(ByVal ValueType As eValueType)
            Me.New(ValueType, False)
        End Sub

        '''<summary>Initialize the statistics calculator with the given type of elements.</summary>
        '''<param name="ValueType">Type of elements (linear, log, ...).</param>
        Public Sub New(ByVal ValueType As eValueType, ByVal StoreRawValues As Boolean)
            IntProps.ValueType = ValueType
            Me.StoreRawValues = StoreRawValues
            Clear()
        End Sub

        '''<summary>Clear the data content.</summary>
        Public Sub Clear()
            With IntProps
                .ValueCount = 0
                .InvalidValueCount = 0
                .Maximum = Double.MinValue
                .Minimum = Double.MaxValue
                .Sum = 0
                .SquareSum = 0
            End With
            If Me.StoreRawValues = True Then RawValueStorage = New List(Of Double)
            Sorted = Nothing
        End Sub

        '''<summary>Lineare variance.</summary>
        '''<returns>Lineare variance.</returns>
        Private Function LinVariance() As Double
            Return (IntProps.SquareSum - ((IntProps.Sum * IntProps.Sum) / IntProps.ValueCount)) / (IntProps.ValueCount - 1)
        End Function

        '''<summary>X axis for the selected vector of statistics.</summary>
        Public Shared Function GetAspectVectorXAxis(ByRef Stats() As cSingleValueStatistics) As Double()
            Dim RetVal(Stats.GetUpperBound(0)) As Double
            For Idx As Integer = 0 To RetVal.GetUpperBound(0)
                RetVal(Idx) = Idx
            Next Idx
            Return RetVal
        End Function

        Public Shared Function GetAspectVector(ByRef Stats() As cSingleValueStatistics, ByVal Aspect As eAspects) As Double()
            Dim RetVal(Stats.GetUpperBound(0)) As Double
            Select Case Aspect
                Case eAspects.Mean
                    For Idx As Integer = 0 To RetVal.GetUpperBound(0)
                        RetVal(Idx) = Stats(Idx).Mean
                    Next Idx
                Case eAspects.Maximum
                    For Idx As Integer = 0 To RetVal.GetUpperBound(0)
                        RetVal(Idx) = Stats(Idx).Maximum
                    Next Idx
                Case eAspects.Minimum
                    For Idx As Integer = 0 To RetVal.GetUpperBound(0)
                        RetVal(Idx) = Stats(Idx).Minimum
                    Next Idx
                Case eAspects.Sigma
                    For Idx As Integer = 0 To RetVal.GetUpperBound(0)
                        RetVal(Idx) = Stats(Idx).Sigma
                    Next Idx
            End Select
            Return RetVal
        End Function

    End Class

End Namespace