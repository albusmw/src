Option Explicit On
Option Strict On

'''<summary>The JPL HORIZONS on-line solar system data and ephemeris computation service provides access to key solar system data and flexible production of highly accurate ephemerides for solar system objects (840636 asteroids, 3598 comets, 210 planetary satellites, 8 planets, the Sun, L1, L2, select spacecraft, and system barycenters).</summary>
'''<remarks>HORIZONS is provided by the Solar System Dynamics Group of the Jet Propulsion Laboratory.</remarks>
Class cHorizons_Loader

    Private ColSplit As Char = ","c

    '''<summary>Observer quantities that can be queried.</summary>
    Public Structure sQuantities
        Public Const Astrometric_RA_DEC As String = "1"
        Public Const Apparent_RA_DEC As String = "2"
        Public Const Apparent_AZ_EL As String = "4"
        Public Const VisualMag_SurfaceBrght As String = "9"
        Public Const IlluminatedFraction As String = "10"
        Public Const TargetAngularDiameter As String = "13"
        Public Const ObserverSubLong_SubLat As String = "14"
        Public Const HeliocentricRange_RangeRate As String = "19"
        Public Const ObserverRange_RangeRate As String = "20"
        Public Const OneWay_DownLeg_LightTime As String = "21"
        Public Const SunObserverTarget_ELONGAngle As String = "23"
        Public Const SunTargetObserver_PHASEAngle As String = "24"
        Public Const Target_Observer_MoonAngle_Illum As String = "25"
    End Structure

    '''<summary>Observer quantities that may be returned.</summary>
    Public Structure sColumns
        '''<summary>Solar presence - '*' for daylight, 'C','N','A' for twilight/dawn, '' else.</summary>
        Public Const SOLAR_PRESENCE As String = "SOLAR_PRESENCE"
        '''<summary>Lunar presence - 'm' for present, '' else.</summary>
        Public Const LUNAR_PRESENCE As String = "LUNAR_PRESENCE"
        '''<summary>Date and time- TODO: check format string.</summary>
        Public Const [Date] As String = "Date__(UT)__HR:MN:SS"
        '''<summary>Apparent RA, airless, HMS/DMS format.</summary>
        Public Const RA__a_app As String = "R.A.__(a-app)"
        '''<summary>Apparent DEC, airless, HMS/DMS format.</summary>
        Public Const DEC__a_app As String = "DEC_(a-app)"
        '''<summary>Apparent azimuth, airless, degree.</summary>
        Public Const AZI__a_app As String = "Azi_(a-app)"
        '''<summary>Apparent elevation, airless, degree.</summary>
        Public Const Elev__a_app As String = "Elev_(a-app)"
    End Structure

    Public Enum eBodies As ULong
        Sun = 10
        Mercury = 199
        Venus = 299
        Moon = 301
        Earth_GeoCenter = 399
        Mars = 499
        Jupiter = 599
        Saturn = 699
        Uranus = 799
        Neptune = 899
        Pluto = 999
    End Enum

    '''<summary>Start date for the query.</summary>
    Public Property FromDate As Date = Now
    '''<summary>End date for the query.</summary>
    Public Property ToDate As Date = Now.AddDays(10)

    '''<summary>Last raw answer as directly generated from the the HORIZONS system.</summary>
    Public ReadOnly Property LastRawAnswer As String()
        Get
            Return MyLastRawAnswer
        End Get
    End Property
    Private MyLastRawAnswer As String() = {}

    '''<summary>All headers generated from the system.</summary>
    Public ReadOnly Property Headers As Dictionary(Of String, Integer)
        Get
            Return MyHeaders
        End Get
    End Property
    Private MyHeaders As Dictionary(Of String, Integer)

    Public Function Go(ByVal Body As eBodies, ByVal StepMinutes As Integer) As List(Of String)
        Return Go(Val(Body).ToString.Trim, StepMinutes)
    End Function

    '''<summary>Run the HORIZONS query.</summary>
    '''<param name="Body">Body number or string to query data from.</param>
    '''<param name="StepMinutes">Stepping [minutes].</param>
    '''<returns></returns>
    '''<seealso cref="https://ssd.jpl.nasa.gov/?horizons_doc"/>
    Public Function Go(ByVal Body As String, ByVal StepMinutes As Integer) As List(Of String)

        'Interface to configure request: https://ssd.jpl.nasa.gov/horizons.cgi
        'See https://ssd.jpl.nasa.gov/horizons_batch.cgi for details

        Dim Example1 As String = "https://ssd.jpl.nasa.gov/horizons_batch.cgi?batch=1&COMMAND='499'&MAKE_EPHEM='YES'&TABLE_TYPE='OBSERVER'&START_TIME='2000-01-01'&STOP_TIME='2000-12-31'&STEP_SIZE='15%20d'&QUANTITIES='1,9,20,23,24'&CSV_FORMAT='YES'"

        'Quantities to query
        Dim Quantities As New List(Of String)
        Quantities.Add(sQuantities.Astrometric_RA_DEC)
        Quantities.Add(sQuantities.Apparent_RA_DEC)
        Quantities.Add(sQuantities.Apparent_AZ_EL)
        Quantities.Add(sQuantities.VisualMag_SurfaceBrght)
        Quantities.Add(sQuantities.IlluminatedFraction)
        Quantities.Add(sQuantities.TargetAngularDiameter)
        Quantities.Add(sQuantities.ObserverSubLong_SubLat)
        Quantities.Add(sQuantities.HeliocentricRange_RangeRate)
        Quantities.Add(sQuantities.ObserverRange_RangeRate)
        Quantities.Add(sQuantities.OneWay_DownLeg_LightTime)
        Quantities.Add(sQuantities.SunObserverTarget_ELONGAngle)
        Quantities.Add(sQuantities.SunTargetObserver_PHASEAngle)
        Quantities.Add(sQuantities.Target_Observer_MoonAngle_Illum)

        'Compose request data
        Dim BatchData As New List(Of String)
        BatchData.Add("https://ssd.jpl.nasa.gov/horizons_batch.cgi?batch=1")
        BatchData.Add("COMMAND='" & Body.ToString.Trim & "'")
        BatchData.Add("MAKE_EPHEM='YES'")                                       '...
        BatchData.Add("TABLE_TYPE='OBSERVER'")
        BatchData.Add("COORD_TYPE='GEODETIC'")                                  '...
        BatchData.Add("CENTER='coord@399'")
        BatchData.Add("SITE_COORD='11.691474,47.878510,0.677'")                 '1st: +=E, -=W; 2nd: +=N, -=S; 3rd: height [km]
        BatchData.Add("START_TIME='" & Format(FromDate, "yyyy-MM-dd") & "'")
        BatchData.Add("STOP_TIME='" & Format(ToDate, "yyyy-MM-dd") & "'")
        BatchData.Add("STEP_SIZE='" & StepMinutes.ToString.Trim & " m'")                                        'h:hour / d:day / m:minutes

        BatchData.Add("CAL_FORMAT= 'CAL'")
        BatchData.Add("TIME_DIGITS= 'SECONDS'")
        BatchData.Add("ANG_FORMAT= 'HMS'")
        BatchData.Add("OUT_UNITS='KM-S'")
        BatchData.Add("RANGE_UNITS= 'AU'")
        BatchData.Add("APPARENT= 'AIRLESS'")
        BatchData.Add("SUPPRESS_RANGE_RATE= 'NO'")
        BatchData.Add("QUANTITIES= '" & Join(Quantities.ToArray, ",") & "'")

        BatchData.Add("SKIP_DAYLT= 'NO'")                                       'set to "YES" to output only during daylight
        BatchData.Add("EXTRA_PREC= 'NO'")
        BatchData.Add("R_T_S_ONLY= 'NO'")
        BatchData.Add("REF_SYSTEM= 'J2000'")
        BatchData.Add("CSV_FORMAT='YES'")
        BatchData.Add("OBJ_DATA='YES'")

        'BatchData.Clear()
        'BatchData.Add("https://ssd.jpl.nasa.gov/horizons_bacth-cgi?batch=1")
        'BatchData.AddRange(Split(My.Resources.ExampleString, System.Environment.NewLine))

        'Create HTML formated request
        Dim Formated As New List(Of String)
        For Each Entry As String In BatchData
            Formated.Add(Entry.Replace("= ", "=").Replace(" ", "%20"))
        Next Entry
        Dim RequestURL As String = Join(Formated.ToArray, "&")

        MyLastRawAnswer = LoadData(RequestURL)

        'Get data
        Dim InData As Boolean = False
        MyHeaders = New Dictionary(Of String, Integer)
        Dim DataToExtract As New List(Of String)
        For Idx As Integer = 0 To MyLastRawAnswer.GetUpperBound(0)
            Dim Line As String = MyLastRawAnswer(Idx)
            If Line.StartsWith("$$SOE") Then
                'Generate the header; columns with no header get no entry in the dictionary
                Dim ColIdx As Integer = 0
                InData = True
                For Each Entry As String In Split(MyLastRawAnswer(Idx - 2), ColSplit)
                    Dim EntryToAdd As String = Entry.Trim
                    If EntryToAdd.Length = 0 Then
                        Select Case ColIdx
                            Case 1
                                'Solar presence is not names and follows the date column
                                MyHeaders.Add(sColumns.SOLAR_PRESENCE, ColIdx)
                            Case 2
                                'Lunar presence is not names and follows the solar presence column
                                MyHeaders.Add(sColumns.LUNAR_PRESENCE, ColIdx)
                        End Select
                    Else
                        If MyHeaders.ContainsKey(EntryToAdd) = False Then MyHeaders.Add(EntryToAdd, ColIdx)
                    End If
                    ColIdx += 1
                Next Entry
            Else
                If Line.StartsWith("$$EOE") Then Exit For
                If InData = True Then
                    Dim Entries As String() = Split(Line, ",")
                    Dim OneLine As New List(Of String)
                    OneLine.Add(Entries(MyHeaders((sColumns.Date))).Trim)
                    OneLine.Add(Entries(MyHeaders((sColumns.SOLAR_PRESENCE))).Trim)
                    OneLine.Add(Entries(MyHeaders((sColumns.LUNAR_PRESENCE))).Trim)
                    OneLine.Add(Entries(MyHeaders((sColumns.Elev__a_app))).Trim.Replace(".", ","))
                    DataToExtract.Add(Join(OneLine.ToArray, "|"))
                End If
            End If
        Next Idx

        Return DataToExtract

    End Function

    Private Shared Function LoadData(ByVal RequestURL As String) As String()

        Dim Downloader As New System.Net.WebClient()
        Downloader.Encoding = System.Text.Encoding.UTF8
        Dim Data As Byte() = Downloader.DownloadData(RequestURL)
        Return Downloader.Encoding.GetString(Data).Split(Chr(10))

    End Function

End Class
