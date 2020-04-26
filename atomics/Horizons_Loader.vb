Option Explicit On
Option Strict On
Imports AstroCalc.NET.Common

'''<summary>The JPL HORIZONS on-line solar system data and ephemeris computation service provides access to key solar system data and flexible production of highly accurate ephemerides for solar system objects (840636 asteroids, 3598 comets, 210 planetary satellites, 8 planets, the Sun, L1, L2, select spacecraft, and system barycenters).</summary>
'''<remarks>HORIZONS is provided by the Solar System Dynamics Group of the Jet Propulsion Laboratory.</remarks>
Class Horizons_Loader

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

    Public Shared Sub Go()

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
        BatchData.Add("COMMAND='301'")                                      '499: Mars, 10=Sun, 301=Moon, 599=Jupiter, 699=Saturn, 799=Uranus, 199=Mercury, 299=Venus
        BatchData.Add("MAKE_EPHEM='YES'")                                   '...
        BatchData.Add("TABLE_TYPE='OBSERVER'")
        BatchData.Add("COORD_TYPE='GEODETIC'")                              '...
        BatchData.Add("CENTER='coord@399'")
        BatchData.Add("SITE_COORD='11.691474,47.878510,0.677'")             '1st: +=E, -=W; 2nd: +=N, -=S; 3rd: height [km]
        BatchData.Add("START_TIME='2018-01-01'")
        BatchData.Add("STOP_TIME='2019-01-01'")
        BatchData.Add("STEP_SIZE='1 h'")                                    'h:hour / d:day / m:minutes

        BatchData.Add("CAL_FORMAT= 'CAL'")
        BatchData.Add("TIME_DIGITS= 'MINUTES'")
        BatchData.Add("ANG_FORMAT= 'HMS'")
        BatchData.Add("OUT_UNITS='KM-S'")
        BatchData.Add("RANGE_UNITS= 'AU'")
        BatchData.Add("APPARENT= 'AIRLESS'")
        BatchData.Add("SUPPRESS_RANGE_RATE= 'NO'")
        BatchData.Add("QUANTITIES= '" & Join(Quantities.ToArray, ",") & "'")

        BatchData.Add("SKIP_DAYLT= 'NO'")                                   'set to "YES" to output only during daylight
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

        Dim Data() As String = LoadData(RequestURL)

        'Get data
        Dim InData As Boolean = False
        Dim Header As String() = {}
        Dim DataToExtract As New List(Of String)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            Dim Line As String = Data(Idx)
            If Line.StartsWith("$$SOE") Then
                InData = True
                Header = Split(Data(Idx - 2), ",")
            Else
                If Line.StartsWith("$$EOE") Then Exit For
                If InData = True Then
                    Dim Entries As String() = Split(Line, ",")
                    DataToExtract.Add(Entries(12).Trim.Replace(".", ","))
                End If
            End If
        Next Idx

        Windows.Forms.Clipboard.Clear()
        Windows.Forms.Clipboard.SetText(Join(DataToExtract.ToArray, System.Environment.NewLine))

        MsgBox("DONE!")

    End Sub

    Private Shared Function LoadData(ByVal RequestURL As String) As String()

        'Net.ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate()     ' <- not required
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls

        'Query data from request URL
        Dim address As New Uri(RequestURL)
        Dim mWC As New System.Net.WebClient
        Dim InStream As IO.Stream = mWC.OpenRead(address)
        Dim InReader As New IO.StreamReader(InStream)
        Dim Data As String() = InReader.ReadToEnd.Split(Chr(10))

        Return Data

    End Function

    Private Function ValidateRemoteCertificate(sender As Object, cert As Security.Cryptography.X509Certificates.X509Certificate, chain As System.Security.Cryptography.X509Certificates.X509Chain, Err As Net.Security.SslPolicyErrors) As Boolean
        If Err = System.Net.Security.SslPolicyErrors.None Then Return True
        Console.WriteLine("X509Certificate [{0}] Policy Error: '{1}'", cert.Subject, Err.ToString())
        Return False
    End Function

End Class
