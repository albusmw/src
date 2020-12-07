Option Explicit On
Option Strict On

'List of named stars:
'http://www.pas.rochester.edu/~emamajek/WGSN/IAU-CSN.txt

Public Class cAstroCats

    'Needs references:  cDownloader.vb
    '                   GZIP.vb

    Public Interface IGenerics
        Property Star As sGeneric
    End Interface

    Public Structure sGeneric

        Public Enum eCatType
            HD
        End Enum

        Public Structure sCat
            Public Value As String
            Public Type As eCatType
            Public Sub New(ByVal NewValue As String, ByVal NewType As eCatType)
                Me.Value = NewValue
                Me.Type = NewType
            End Sub
        End Structure

        Public Enum eMagnitudeType
            Visual
        End Enum

        Public Structure sMagnitude
            Public Value As Double
            Public Type As eMagnitudeType
            Public Sub New(ByVal NewValue As Double, ByVal NewType As eMagnitudeType)
                Me.Value = NewValue
                Me.Type = NewType
            End Sub
        End Structure

        '''<summary>Right ascension [degree].</summary>
        Public RightAscension As Double
        '''<summary>Declination [degree].</summary>
        Public Declination As Double
        '''<summary>Magnitude information.</summary>
        Public Magnitude As sMagnitude
        '''<summary>Catalog information.</summary>
        Public Cat As sCat

        '''<summary>Set the position of the object.</summary>
        Public Sub Invalidate()
            Me.RightAscension = Double.NaN
            Me.Declination = Double.NaN
        End Sub

        '''<summary>Set the position of the object.</summary>
        Public Sub New(ByVal RA As Double, ByVal DE As Double)
            Me.RightAscension = RA
            Me.Declination = DE
        End Sub

        '''<summary>Indicate that the object is empty.</summary>
        Public Function IsNothing() As Boolean
            If Double.IsNaN(RightAscension) = True Then Return True
            If Double.IsNaN(Declination) = True Then Return True
            Return False
        End Function

    End Structure

    '''<summary>Hipparcos Catalogue.</summary>
    Public Class cHipparcos

        '''<summary>Root URL where to load the content from.</summary>
        Public Property RootURL() As String
            Get
                Return MyRootURL
            End Get
            Set(value As String)
                MyRootURL = value
            End Set
        End Property
        Private MyRootURL As String = "ftp://cdsarc.u-strasbg.fr/pub/cats/I/239/hip_main.dat.gz"

        '''<summary>Original file name (as present on server).</summary>
        Public ReadOnly Property OrigFileName() As String
            Get
                Return "hip_main.dat.gz"
            End Get
        End Property

        '''<summary>Extracted file name (as present on local disc after extract).</summary>
        Public ReadOnly Property ExtractedFile() As String
            Get
                Return "hip_main.dat"
            End Get
        End Property

        '''<summary>Local folder where to find the catalog file.</summary>
        Public ReadOnly Property LocalFolder() As String = String.Empty

        '''<summary>Downloader for content.</summary>
        Private MyDownloader As New cDownloader

        Public Event Currently(ByVal Info As String)

        Public Catalog As New Dictionary(Of Integer, sHIPEntry)

        'Hipparcos readme:
        ' ftp://cdsarc.u-strasbg.fr/pub/cats/I/239/ReadMe


        Public Structure sHIPEntry : Implements IGenerics
            '''<summary>Generic information.</summary>
            Public Property Star As sGeneric Implements IGenerics.Star
            '''<summary>Henry Draper Catalog (HD) number.</summary>
            Public HIP As Integer
            '''<summary>Magnitude.</summary>
            Public Vmag As Single
        End Structure

        Public Sub Download()
            MyDownloader.InitWebClient()
            RaiseEvent Currently("Loading Hipparcos Catalogue data from VIZIER FTP ...")
            MyDownloader.DownloadFile(RootURL, LocalFolder & "\" & OrigFileName)
            GZIP.DecompressTo(LocalFolder & "\" & OrigFileName, LocalFolder & "\" & ExtractedFile)
            System.IO.File.Delete(LocalFolder & "\" & OrigFileName)
            RaiseEvent Currently("   DONE.")
        End Sub

        Public Sub New()
            Me.New(String.Empty)
        End Sub

        Public Sub New(ByVal CatalogFolder As String)

            Catalog = New Dictionary(Of Integer, sHIPEntry)
            Dim ErrorCount As Integer = 0

            LocalFolder = CatalogFolder
            If System.IO.File.Exists(LocalFolder & "\" & ExtractedFile) = False Then Exit Sub

            'TODO: This is not the finalized code ...

            For Each BlockContent As String In System.IO.File.ReadAllLines(LocalFolder & "\" & ExtractedFile, System.Text.Encoding.ASCII)
                'Add some spaces in the back to avoid problems during parsing
                BlockContent &= "                           "
                Try
                    Dim NewEntry As New sHIPEntry With {
                        .HIP = CInt(BlockContent.Substring(8, 6).Trim)
                    }
                    Dim RA As Double = (CInt(BlockContent.Substring(18, 2)) + (CInt(BlockContent.Substring(20, 3)) / 600)) * 15
                    Dim DE As Double = (CInt(BlockContent.Substring(24, 2)) + (CInt(BlockContent.Substring(26, 2)) / 60)) * CDbl(IIf(BlockContent.Substring(23, 1) = "-", -1, 1))
                    NewEntry.Star = New sGeneric(RA, DE)
                    NewEntry.Vmag = CSng(Val(BlockContent.Substring(29, 5)))
                    Catalog.Add(NewEntry.HIP, NewEntry)
                Catch ex As Exception
                    ErrorCount += 1
                End Try
            Next BlockContent

        End Sub

    End Class

    '''<summary>Simbad database.</summary>
    '''<remarks>The SIMBAD astronomical database. The CDS reference database for astronomical objects.</remarks>
    Public Class cSimbad

        'Documentation:
        ' http://simbad.u-strasbg.fr/Pages/guide/sim-url.htx

        Public Class Descriptors
            Public Shared Identifier As String = "Identifier"
            Public Shared Typ As String = "Typ"
            Public Shared Typ_verbose As String = "Typ_verbose"
            Public Shared Coord_RA As String = "Coord_RA"
            Public Shared Coord_DEC As String = "Coord_DEC"
            Public Shared Mag_V As String = "Mag_V"
        End Class

        'Example query:
        'Criteria=Vmag%3c4&submit=submit%20query&OutputMode=LIST&maxObject=1000&CriteriaFile=&output.format=ASCII

        '''<summary>Root URL where to load the content from.</summary>
        Public Shared Property RootURL_SimSam() As String = "http://simbad.u-strasbg.fr/simbad/sim-sam?"
        '''<summary>Root URL where to load the content from.</summary>
        Public Shared Property RootURL_SimId() As String = "http://simbad.u-strasbg.fr/simbad/sim-id?"

        '''<summary>Downloader for content.</summary>
        Private Shared MyDownloader As New cDownloader

        '''<summary>Run a specific query against the SIMBAD database.</summary>
        Public Shared Function QuerySimSam(ByVal QueryToRun As String) As String
            MyDownloader.InitWebClient()
            Return MyDownloader.DownloadString(RootURL_SimSam & QueryToRun)
        End Function

        '''<summary>Run a specific query against the SIMBAD database.</summary>
        Public Shared Function QuerySimID(ByVal QueryToRun As String) As String
            MyDownloader.InitWebClient()
            Return MyDownloader.DownloadString(RootURL_SimId & QueryToRun)
        End Function

        '''<summary>Run a specific conplete catalog on the SIMBAD database.</summary>
        Public Shared Function QuerySimID_Cat(ByVal Catalog As String) As String
            MyDownloader.InitWebClient()
            Dim Query As String = RootURL_SimId & "Ident=" & Catalog & "&NbIdent=cat&submit=submit+id&output.format=ASCII"
            Return MyDownloader.DownloadString(Query)
        End Function

        '''<summary>Get the object data of a specific named element.</summary>
        Public Shared Function GetNamedElement(ByVal ElementName As String, ByRef RA As Double, ByRef Dec As Double) As String

            RA = Double.NaN
            Dec = Double.NaN

            Try
                Dim req As Net.HttpWebRequest = CType(Net.WebRequest.Create(RootURL_SimId & "Ident=" & ElementName & "&output.format=ASCII"), Net.HttpWebRequest)
                req.Method = "Get"
                Dim Answer As String = (New IO.StreamReader(CType(req.GetResponse(), Net.HttpWebResponse).GetResponseStream)).ReadToEnd
                Return Answer
            Catch ex As Exception
                Return String.Empty
            End Try

        End Function

        '''<summary>Parse the returned SIMBAD answer.</summary>
        Public Shared Function ParseAnswer(ByVal SIMBADAnswer As String) As List(Of Dictionary(Of String, Object))
            Dim RetVal As New List(Of Dictionary(Of String, Object))
            Dim InDataCounter As Integer = 0                        'there are 2 starting lines which are ignored (header and -------)
            Dim Header As New Dictionary(Of String, Integer)        'mapping between header and column index
            For Each Line As String In SIMBADAnswer.Split(Chr(10))
                Dim SingleLine As String() = TrimObjectCode(Split(Line, "|"))
                If SingleLine.Length > 5 Then
                    'Process the header
                    If InDataCounter = 0 Then
                        For Idx As Integer = 0 To SingleLine.GetUpperBound(0)
                            Header.Add(SingleLine(Idx).Replace(" ", String.Empty), Idx)
                        Next Idx
                    End If
                    If InDataCounter >= 2 Then
                        Dim ObjectDescription As New Dictionary(Of String, Object)
                        Dim OType As String = CStr(GetNamedElement(SingleLine, Header, "typ"))
                        Dim Coords As String = CStr(GetNamedElement(SingleLine, Header, "coord1 (ICRS,J2000/2000)"))
                        ObjectDescription.Add(Descriptors.Identifier, GetNamedElement(SingleLine, Header, "identifier"))                 'identifier
                        ObjectDescription.Add(Descriptors.Typ, OType)
                        If IsNothing(OType) = False Then ObjectDescription.Add(Descriptors.Typ_verbose, SimbadObjectCode(OType))
                        If IsNothing(Coords) = False Then
                            Dim CoordSplit As String() = Split(Coords, " ")
                            ObjectDescription.Add(Descriptors.Coord_RA, Val(CoordSplit(0)) + (Val(CoordSplit(1)) / 60) + (Val(CoordSplit(2)) / 3600))
                            ObjectDescription.Add(Descriptors.Coord_DEC, Val(CoordSplit(3)) + (Val(CoordSplit(4)) / 60) + (Val(CoordSplit(5)) / 3600))
                        End If
                        ObjectDescription.Add(Descriptors.Mag_V, GetNamedElement(SingleLine, Header, "Mag V"))
                        RetVal.Add(ObjectDescription)
                    End If
                    InDataCounter += 1
                End If
            Next Line
            Return RetVal
        End Function

        ''' <summary>Get the requested element from the SIMBAD answer.</summary>
        ''' <param name="SplitAnswer">Splitted single SIMBAD answer line.</param>
        ''' <param name="Header">Detected header elements.</param>
        ''' <param name="Element">Requested element.</param>
        Public Shared Function GetNamedElement(ByRef SplitAnswer As String(), ByRef Header As Dictionary(Of String, Integer), ByVal Element As String) As Object
            Element = Element.Replace(" ", String.Empty)
            If Header.ContainsKey(Element) = True Then
                Dim Idx As Integer = Header(Element)
                If SplitAnswer.GetUpperBound(0) >= Idx Then
                    Return SplitAnswer(Idx)
                Else
                    Return Nothing                              'answer is not long enough
                End If
            Else
                Return Nothing                                  'requested element is not present
            End If
        End Function

        '''<summary>Translate the code from webpage codes webpage to VB code (which can be found below ...).</summary>
        '''<remarks>http://simbad.u-strasbg.fr/simbad/sim-display?data=otypes&option=display+numeric+codes</remarks>
        Public Shared Function GenObjectCode(ByVal TextToConvert As String) As String
            Dim Code As New List(Of String)
            Code.Add("Public Shared ReadOnly SimbadObjectCode As New Dictionary(Of String, String) From {")
            For Each Line As String In TextToConvert.Split(Chr(10))
                Dim SingleLine As String() = TrimObjectCode(Split(Line, vbTab))
                If SingleLine.Length >= 4 Then
                    Code.Add("{" & Quote(SingleLine(2)) & ", " & Quote(SingleLine(1) & " (" & SingleLine(3) & ")") & "},")
                End If
            Next Line
            Code.Add("}")
            Return Join(Code.ToArray, System.Environment.NewLine)
        End Function

        '''<summary>Quote the given string.</summary>
        Private Shared Function Quote(ByVal ObjectCode As String) As String
            Return Chr(34) & ObjectCode & Chr(34)
        End Function

        '''<summary>Remove spaces, ... from the object codes.</summary>
        Private Shared Function TrimObjectCode(ByVal SingleLine As String()) As String()
            Dim RetVal As New List(Of String)
            For Each Entry As String In SingleLine
                RetVal.Add(TrimObjectCode(Entry))
            Next Entry
            Return RetVal.ToArray
        End Function

        '''<summary>Remove spaces, ... from the object code.</summary>
        Private Shared Function TrimObjectCode(ByVal ObjectCode As String) As String
            ObjectCode = ObjectCode.TrimStart(New Char() {Chr(183), Chr(32), Chr(10), Chr(13)}).TrimEnd(New Char() {Chr(183), Chr(32), Chr(10), Chr(13)})
            Return ObjectCode
        End Function

        Public Shared Function SortByMagV(A As String, B As String) As Integer
            Return Val(A.Substring(0, 5)).CompareTo(Val(B.Substring(0, 5)))
        End Function

        '''<summary>SIMBAD object codes and description.</summary>
        Public Shared ReadOnly SimbadObjectCode As New Dictionary(Of String, String) From {
{"?", "Unknown (Object Of unknown nature)"},
{"ev", "Transient (transient Event)"},
{"Rad", "Radio (Radio-source)"},
{"mR", "Radio(m) (metric Radio-source)"},
{"cm", "Radio(cm) (centimetric Radio-source)"},
{"mm", "Radio(mm) (millimetric Radio-source)"},
{"smm", "Radio(Sub-mm) (Sub-millimetric source)"},
{"HI", "HI (HI (21cm) source)"},
{"rB", "radioBurst (radio Burst)"},
{"Mas", "Maser (Maser)"},
{"IR", "IR (Infra-Red source)"},
{"FIR", "IR>30um (Far-IR source (λ >= 30 µm))"},
{"NIR", "IR<10um (Near-IR source (λ < 10 µm))"},
{"red", "Red (Very red source)"},
{"ERO", "RedExtreme (Extremely Red Object)"},
{"blu", "Blue (Blue Object)"},
{"UV", "UV (UV-emission source)"},
{"X", "X (X-ray source)"},
{"UX?", "ULX? (Ultra-luminous X-ray candidate)"},
{"ULX", "ULX (Ultra-luminous X-ray source)"},
{"gam", "gamma (gamma-ray source)"},
{"gB", "gammaBurst (gamma-ray Burst)"},
{"err", "Inexistent (Not an Object (Error, artefact, ...))"},
{"grv", "Gravitation (Gravitational Source)"},
{"Lev", "LensingEv ((Micro)Lensing Event)"},
{"LS?", "Candidate_LensSystem (Possible gravitational lens System)"},
{"Le?", "Candidate_Lens (Possible gravitational lens)"},
{"LI?", "Possible_lensImage (Possible gravitationally lensed image)"},
{"gLe", "GravLens (Gravitational Lens)"},
{"gLS", "GravLensSystem (Gravitational Lens System (lens+images))"},
{"GWE", "GravWaveEvent (Gravitational Wave Event)"},
{"..?", "Candidates (Candidate objects)"},
{"G?", "Possible_G (Possible Galaxy)"},
{"SC?", "Possible_SClG (Possible Supercluster Of Galaxies)"},
{"C?G", "Possible_ClG (Possible Cluster Of Galaxies)"},
{"Gr?", "Possible_GrG (Possible Group Of Galaxies)"},
{"As?", "Possible_As* ()"},
{"**?", "Candidate_** (Physical Binary Candidate)"},
{"EB?", "Candidate_EB* (Eclipsing Binary Candidate)"},
{"Sy?", "Candidate_Symb* (Symbiotic Star Candidate)"},
{"CV?", "Candidate_CV* (Cataclysmic Binary Candidate)"},
{"No?", "Candidate_Nova (Nova Candidate)"},
{"XB?", "Candidate_XB* (X-ray binary Candidate)"},
{"LX?", "Candidate_LMXB (Low-Mass X-ray binary Candidate)"},
{"HX?", "Candidate_HMXB (High-Mass X-ray binary Candidate)"},
{"Pec?", "Candidate_Pec* (Possible Peculiar Star)"},
{"Y*?", "Candidate_YSO (Young Stellar Object Candidate)"},
{"pr?", "Candidate_pMS* (Pre-main sequence Star Candidate)"},
{"TT?", "Candidate_TTau* (T Tau star Candidate)"},
{"C*?", "Candidate_C* (Possible Carbon Star)"},
{"S*?", "Candidate_S* (Possible S Star)"},
{"OH?", "Candidate_OH (Possible Star With envelope Of OH/IR type)"},
{"CH?", "Candidate_CH (Possible Star With envelope Of CH type)"},
{"WR?", "Candidate_WR* (Possible Wolf-Rayet Star)"},
{"Be?", "Candidate_Be* (Possible Be Star)"},
{"Ae?", "Candidate_Ae* (Possible Herbig Ae/Be Star)"},
{"HB?", "Candidate_HB* (Possible Horizontal Branch Star)"},
{"RR?", "Candidate_RRLyr (Possible Star Of RR Lyr type)"},
{"Ce?", "Candidate_Cepheid (Possible Cepheid)"},
{"RB?", "Candidate_RGB* (Possible Red Giant Branch star)"},
{"sg?", "Candidate_SG* (Possible Supergiant star)"},
{"s?r", "Candidate_RSG* (Possible Red supergiant star)"},
{"s?y", "Candidate_YSG* (Possible Yellow supergiant star)"},
{"s?b", "Candidate_BSG* (Possible Blue supergiant star)"},
{"AB?", "Candidate_AGB* (Asymptotic Giant Branch Star candidate)"},
{"LP?", "Candidate_LP* (Long Period Variable candidate)"},
{"Mi?", "Candidate_Mi* (Mira candidate)"},
{"sv?", "Candiate_sr* (Semi-regular variable candidate)"},
{"pA?", "Candidate_post-AGB* (Post-AGB Star Candidate)"},
{"BS?", "Candidate_BSS (Candidate blue Straggler Star)"},
{"HS?", "Candidate_Hsd (Hot subdwarf candidate)"},
{"WD?", "Candidate_WD* (White Dwarf Candidate)"},
{"N*?", "Candidate_NS (Neutron Star Candidate)"},
{"BH?", "Candidate_BH (Black Hole Candidate)"},
{"SN?", "Candidate_SN* (SuperNova Candidate)"},
{"LM?", "Candidate_low-mass* (Low-mass star candidate)"},
{"BD?", "Candidate_brownD* (Brown Dwarf Candidate)"},
{"mul", "multiple_object (Composite Object)"},
{"reg", "Region (Region defined In the sky)"},
{"vid", "Void (Underdense region Of the Universe)"},
{"SCG", "SuperClG (Supercluster Of Galaxies)"},
{"ClG", "ClG (Cluster Of Galaxies)"},
{"GrG", "GroupG (Group Of Galaxies)"},
{"CGG", "Compact_Gr_G (Compact Group Of Galaxies)"},
{"PaG", "PairG (Pair Of Galaxies)"},
{"IG", "IG (Interacting Galaxies)"},
{"C?*", "Cl*? (Possible (open) star cluster)"},
{"Gl?", "GlCl? (Possible Globular Cluster)"},
{"Cl*", "Cl* (Cluster Of Stars)"},
{"GlC", "GlCl (Globular Cluster)"},
{"OpC", "OpCl (Open (galactic) Cluster)"},
{"As*", "Assoc* (Association Of Stars)"},
{"St*", "Stream* (Stellar Stream)"},
{"MGr", "MouvGroup (Moving Group)"},
{"**", "** (Double Or multiple star)"},
{"EB*", "EB* (Eclipsing binary)"},
{"Al*", "EB*Algol (Eclipsing binary Of Algol type)"},
{"bL*", "EB*betLyr (Eclipsing binary Of beta Lyr type)"},
{"WU*", "EB*WUMa (Eclipsing binary Of W UMa type)"},
{"EP*", "EB*Planet (Star showing eclipses by its planet)"},
{"SB*", "SB* (Spectroscopic binary)"},
{"El*", "EllipVar (Ellipsoidal variable Star)"},
{"Sy*", "Symbiotic* (Symbiotic Star)"},
{"CV*", "CataclyV* (Cataclysmic Variable Star)"},
{"DQ*", "DQHer (CV DQ Her type (intermediate polar))"},
{"AM*", "AMHer (CV Of AM Her type (polar))"},
{"NL*", "Nova-Like (Nova-Like Star)"},
{"No*", "Nova (Nova)"},
{"DN*", "DwarfNova (Dwarf Nova)"},
{"XB*", "XB (X-ray Binary)"},
{"LXB", "LMXB (Low Mass X-ray Binary)"},
{"HXB", "HMXB (High Mass X-ray Binary)"},
{"ISM", "ISM (Interstellar matter)"},
{"PoC", "PartofCloud (Part Of Cloud)"},
{"PN?", "PN? (Possible Planetary Nebula)"},
{"CGb", "ComGlob (Cometary Globule)"},
{"bub", "Bubble (Bubble)"},
{"EmO", "EmObj (Emission Object)"},
{"Cld", "Cloud (Cloud)"},
{"GNe", "GalNeb (Galactic Nebula)"},
{"BNe", "BrNeb (Bright Nebula)"},
{"DNe", "DkNeb (Dark Cloud (nebula))"},
{"RNe", "RfNeb (Reflection Nebula)"},
{"MoC", "MolCld (Molecular Cloud)"},
{"glb", "Globule (Globule (low-mass dark cloud))"},
{"cOr", "denseCore (Dense core)"},
{"SFR", "SFregion (Star forming region)"},
{"HVC", "HVCld (High-velocity Cloud)"},
{"HII", "HII (HII (ionized) region)"},
{"PN", "PN (Planetary Nebula)"},
{"sh", "HIshell (HI shell)"},
{"SR?", "SNR? (SuperNova Remnant Candidate)"},
{"SNR", "SNR (SuperNova Remnant)"},
{"cir", "Circumstellar (CircumStellar matter)"},
{"Of?", "outflow? (Outflow candidate)"},
{"out", "Outflow (Outflow)"},
{"HH", "HH (Herbig-Haro Object)"},
{"*", "Star (Star)"},
{"*iC", "*inCl (Star In Cluster)"},
{"*In", "*inNeb (Star In Nebula)"},
{"*iA", "*inAssoc (Star In Association)"},
{"*i*", "*In** (Star In Double system)"},
{"V*?", "V*? (Star suspected Of Variability)"},
{"Pe*", "Pec* (Peculiar Star)"},
{"HB*", "HB* (Horizontal Branch Star)"},
{"Y*O", "YSO (Young Stellar Object)"},
{"Ae*", "Ae* (Herbig Ae/Be star)"},
{"Em*", "Em* (Emission-line Star)"},
{"Be*", "Be* (Be Star)"},
{"BS*", "BlueStraggler (Blue Straggler Star)"},
{"RG*", "RGB* (Red Giant Branch star)"},
{"AB*", "AGB* (Asymptotic Giant Branch Star (He-burning))"},
{"C*", "C* (Carbon Star)"},
{"S*", "S* (S Star)"},
{"sg*", "SG* (Evolved supergiant star)"},
{"s*r", "RedSG* (Red supergiant star)"},
{"s*y", "YellowSG* (Yellow supergiant star)"},
{"s*b", "BlueSG* (Blue supergiant star)"},
{"HS*", "HotSubdwarf (Hot subdwarf)"},
{"pA*", "post-AGB* (Post-AGB Star (proto-PN))"},
{"WD*", "WD* (White Dwarf)"},
{"ZZ*", "pulsWD* (Pulsating White Dwarf)"},
{"LM*", "low-mass* (Low-mass star (M<1SolMass))"},
{"BD*", "brownD* (Brown Dwarf (M<0.08solMass))"},
{"N*", "Neutron* (Confirmed Neutron Star)"},
{"OH*", "OH/IR (OH/IR star)"},
{"CH*", "CH (Star With envelope Of CH type)"},
{"pr*", "pMS* (Pre-main sequence Star)"},
{"TT*", "TTau* (T Tau-type Star)"},
{"WR*", "WR* (Wolf-Rayet Star)"},
{"PM*", "PM* (High proper-motion Star)"},
{"HV*", "HV* (High-velocity Star)"},
{"V*", "V* (Variable Star)"},
{"Ir*", "Irregular_V* (Variable Star Of irregular type)"},
{"Or*", "Orion_V* (Variable Star Of Orion Type)"},
{"RI*", "Rapid_Irreg_V* (Variable Star With rapid variations)"},
{"Er*", "Eruptive* (Eruptive variable Star)"},
{"Fl*", "Flare* (Flare Star)"},
{"FU*", "FUOr (Variable Star Of FU Ori type)"},
{"RC*", "Erupt*RCrB (Variable Star Of R CrB type)"},
{"RC?", "RCrB_Candidate (Variable Star Of R CrB type candiate)"},
{"Ro*", "RotV* (Rotationally variable Star)"},
{"a2*", "RotV*alf2CVn (Variable Star Of alpha2 CVn type)"},
{"Psr", "Pulsar (Pulsar)"},
{"BY*", "BYDra (Variable Of BY Dra type)"},
{"RS*", "RSCVn (Variable Of RS CVn type)"},
{"Pu*", "PulsV* (Pulsating variable Star)"},
{"RR*", "RRLyr (Variable Star Of RR Lyr type)"},
{"Ce*", "Cepheid (Cepheid variable Star)"},
{"dS*", "PulsV*delSct (Variable Star Of delta Sct type)"},
{"RV*", "PulsV*RVTau (Variable Star Of RV Tau type)"},
{"WV*", "PulsV*WVir (Variable Star Of W Vir type)"},
{"bC*", "PulsV*bCep (Variable Star Of beta Cep type)"},
{"cC*", "deltaCep (Classical Cepheid (delta Cep type))"},
{"gD*", "gammaDor (Variable Star Of gamma Dor type)"},
{"SX*", "pulsV*SX (Variable Star Of SX Phe type (subdwarf))"},
{"LP*", "LPV* (Long-period variable star)"},
{"Mi*", "Mira (Variable Star Of Mira Cet type)"},
{"sr*", "semi-regV* (Semi-regular pulsating Star)"},
{"SN*", "SN (SuperNova)"},
{"su*", "Sub-stellar (Sub-stellar Object)"},
{"Pl?", "Planet? (Extra-solar Planet Candidate)"},
{"Pl", "Planet (Extra-solar Confirmed Planet)"},
{"G", "Galaxy (Galaxy)"},
{"PoG", "PartofG (Part Of a Galaxy)"},
{"GiC", "GinCl (Galaxy In Cluster Of Galaxies)"},
{"BiC", "BClG (Brightest galaxy In a Cluster (BCG))"},
{"GiG", "GinGroup (Galaxy In Group Of Galaxies)"},
{"GiP", "GinPair (Galaxy In Pair Of Galaxies)"},
{"HzG", "High_z_G (Galaxy With high redshift)"},
{"ALS", "AbsLineSystem (Absorption Line system)"},
{"LyA", "Ly-alpha_ALS (Ly alpha Absorption Line system)"},
{"DLA", "DLy-alpha_ALS (Damped Ly-alpha Absorption Line system)"},
{"mAL", "metal_ALS (metallic Absorption Line system)"},
{"LLS", "Ly-limit_ALS (Lyman limit system)"},
{"BAL", "Broad_ALS (Broad Absorption Line system)"},
{"rG", "RadioG (Radio Galaxy)"},
{"H2G", "HII_G (HII Galaxy)"},
{"LSB", "LSB_G (Low Surface Brightness Galaxy)"},
{"AG?", "AGN_Candidate (Possible Active Galaxy Nucleus)"},
{"Q?", "QSO_Candidate (Possible Quasar)"},
{"Bz?", "Blazar_Candidate (Possible Blazar)"},
{"BL?", "BLLac_Candidate (Possible BL Lac)"},
{"EmG", "EmG (Emission-line galaxy)"},
{"SBG", "StarburstG (Starburst Galaxy)"},
{"bCG", "BlueCompG (Blue compact Galaxy)"},
{"LeI", "LensedImage (Gravitationally Lensed Image)"},
{"LeG", "LensedG (Gravitationally Lensed Image Of a Galaxy)"},
{"LeQ", "LensedQ (Gravitationally Lensed Image Of a Quasar)"},
{"AGN", "AGN (Active Galaxy Nucleus)"},
{"LIN", "LINER (LINER-type Active Galaxy Nucleus)"},
{"SyG", "Seyfert (Seyfert Galaxy)"},
{"Sy1", "Seyfert_1 (Seyfert 1 Galaxy)"},
{"Sy2", "Seyfert_2 (Seyfert 2 Galaxy)"},
{"Bla", "Blazar (Blazar)"},
{"BLL", "BLLac (BL Lac - type Object)"},
{"OVV", "OVV (Optically Violently Variable Object)"},
{"QSO", "QSO (Quasar)"}
}

    End Class

    Public Class cHenryDraper

        '''<summary>Root URL where to load the content from.</summary>
        Public Property RootURL() As String = "http://cdsarc.u-strasbg.fr/vizier/ftp/cats/III/135A/catalog.dat.gz"

        '''<summary>Downloader for content.</summary>
        Private MyDownloader As New cDownloader

        Public Event Currently(ByVal Info As String)

        Public Catalog As New Dictionary(Of Integer, sHDEEntry)

        'Henry Draper Catalogue
        ' Download: http://cdsarc.u-strasbg.fr/cgi-bin/qcat?III/135A

        'Byte-per-byte Description of file: catalog.dat
        '--------------------------------------------------------------------------------
        '   Bytes Format  Units   Label  Explanations
        '--------------------------------------------------------------------------------
        '   1-  6  I6     ---     HD     [1/272150]+ Henry Draper Catalog (HD) number
        '   7- 18  A12    ---     DM     Durchmusterung identification (1)
        '  19- 20  I2     h       RAh    Hours RA, equinox B1900, epoch 1900.0
        '  21- 23  I3     0.1min  RAdm   Deci-Minutes RA, equinox B1900, epoch 1900.0
        '      24  A1     ---     DE-    Sign Dec, equinox B1900, epoch 1900.0
        '  25- 26  I2     deg     DEd    Degrees Dec, equinox B1900, epoch 1900.0
        '  27- 28  I2     arcmin  DEm    Minutes Dec, equinox B1900, epoch 1900.0
        '      29  I1     ---   q_Ptm    [0/1]? Code for Ptm: 0 = measured, 1 = value
        '                                       inferred from Ptg and spectral type
        '  30- 34  F5.2   mag     Ptm    ? Photovisual magnitude (2)
        '      35  A1     ---   n_Ptm    [C] 'C' if Ptm is combined value with Ptg
        '      36  I1     ---   q_Ptg    [0/1]? Code for Ptg: 0 = measured, 1 = value
        '                                       inferred from Ptm and spectral type
        '  37- 41  F5.2   mag     Ptg    ? Photographic magnitude (2)
        '      42  A1     ---   n_Ptg    [C] 'C' if Ptg is combined value for this
        '                                  entry and the following or preceding entry
        '  43- 45  A3     ---     SpT    Spectral type
        '  46- 47  A2     ---     Int    [ 0-9B] Photographic intensity of spectrum (3)
        '      48  A1     ---     Rem    [DEGMR*] Remarks, see note (4)
        '--------------------------------------------------------------------------------

        Public Structure sHDEEntry : Implements IGenerics
            '''<summary>Generic information.</summary>
            Public Property Star As sGeneric Implements IGenerics.Star
            '''<summary>Henry Draper Catalog (HD) number.</summary>
            Public HD As Integer
            '''<summary>Durchmusterung identification.</summary>
            Public DM As String
            '''<summary>Photovisual magnitude.</summary>
            Public MagnitudePhotovisual As Single
            '''<summary>Photographic magnitude.</summary>
            Public MagnitudePhotographic As Single
        End Structure

        '''<summary>Original file name (as present on server).</summary>
        Public ReadOnly Property OrigFileName() As String
            Get
                Return "catalog.dat.gz"
            End Get
        End Property

        '''<summary>Extracted file name (as present on local disc after extract).</summary>
        Public ReadOnly Property ExtractedFile() As String
            Get
                Return "catalog.dat"
            End Get
        End Property

        '''<summary>Local folder where to find the catalog file.</summary>
        Public ReadOnly Property LocalFolder() As String = String.Empty

        Public Sub Download()
            MyDownloader.InitWebClient()
            RaiseEvent Currently("Loading Henry Draper Catalogue data from VIZIER FTP ...")
            MyDownloader.DownloadFile(RootURL, LocalFolder & "\" & OrigFileName)
            GZIP.DecompressTo(LocalFolder & "\" & OrigFileName, LocalFolder & "\" & ExtractedFile)
            System.IO.File.Delete(LocalFolder & "\" & OrigFileName)
            RaiseEvent Currently("   DONE.")
        End Sub


        Public Sub New()
            Me.New(String.Empty)
        End Sub

        Public Sub New(ByVal CatalogFolder As String)

            Catalog = New Dictionary(Of Integer, sHDEEntry)
            Dim ErrorCount As Integer = 0

            LocalFolder = CatalogFolder
            If System.IO.File.Exists(LocalFolder & "\" & ExtractedFile) = False Then Exit Sub

            For Each BlockContent As String In System.IO.File.ReadAllLines(LocalFolder & "\" & ExtractedFile, System.Text.Encoding.ASCII)
                'Add some spaces in the back to avoid problems during parsing
                BlockContent &= "                           "
                Try
                    Dim NewEntry As New sHDEEntry With {
                        .HD = CInt(BlockContent.Substring(0, 6).Trim),
                        .DM = BlockContent.Substring(6, 12)
                    }
                    Dim RA As Double = (CInt(BlockContent.Substring(18, 2)) + (CInt(BlockContent.Substring(20, 3)) / 600)) * 15
                    Dim DE As Double = (CInt(BlockContent.Substring(24, 2)) + (CInt(BlockContent.Substring(26, 2)) / 60)) * CDbl(IIf(BlockContent.Substring(23, 1) = "-", -1, 1))
                    NewEntry.Star = New sGeneric(RA, DE)
                    NewEntry.MagnitudePhotovisual = CSng(Val(BlockContent.Substring(29, 5)))
                    NewEntry.MagnitudePhotographic = CSng(Val(BlockContent.Substring(36, 5)))
                    Catalog.Add(NewEntry.HD, NewEntry)
                Catch ex As Exception
                    ErrorCount += 1
                End Try
            Next BlockContent

        End Sub

    End Class

    Public Class cNGC

        Public Catalog As List(Of sNGCEntry)

        'NGC Catalogue
        ' Download: ftp://cdsarc.u-strasbg.fr/pub/cats/VII/118

        'Byte-per-byte Description of file: ngc2000.dat
        '--------------------------------------------------------------------------------
        '   Bytes Format  Units   Label    Explanations
        '--------------------------------------------------------------------------------
        '   1-  5  A5     ---     Name     NGC or IC designation (preceded by I)
        '   7-  9  A3     ---     Type     Object classification (1)
        '  11- 12  I2     h       RAh      Right Ascension B2000 (hours)
        '  14- 17  F4.1   min     RAm      Right Ascension B2000 (minutes)
        '      20  A1     ---     DE-      Declination B2000 (sign)
        '  21- 22  I2     deg     DEd      Declination B2000 (degrees)
        '  24- 25  I2     arcmin  DEm      Declination B2000 (minutes)
        '      27  A1     ---     Source   Source of entry (2)
        '  30- 32  A3     ---     Const    Constellation
        '      33  A1     ---     l_size   [<] Limit on Size
        '  34- 38  F5.1   arcmin  size     ? Largest dimension
        '  41- 44  F4.1   mag     mag      ? Integrated magnitude, visual or photographic
        '                                      (see n_mag)
        '      45  A1     ---     n_mag    [p] 'p' if mag is photographic (blue)
        '  47- 96  A50    ---     Desc     Description of the object (3)
        '--------------------------------------------------------------------------------

        Public Structure sNGCEntry : Implements IGenerics
            '''<summary>NGC or IC designation (preceded by I).</summary>
            Public Name As String
            '''<summary>Object classification.</summary>
            Public Classification As String
            '''<summary>Generic information.</summary>
            Public Property Star As sGeneric Implements IGenerics.Star
            '''<summary>Constellation name.</summary>
            Public Constellation As String
            '''<summary>Largest dimension [arcmin].</summary>
            Public Dimension As Single
            '''<summary>Photovisual magnitude.</summary>
            Public Magnitude As Single
            '''<summary>Description of the object.</summary>
            Public Description As String
            '''<summary>Decide if the element is empty.</summary>
            Public Function IsNothing() As Boolean
                Return Star.IsNothing
            End Function
            Public Shared Function Empty() As sNGCEntry
                Dim RetVal As New sNGCEntry
                With RetVal
                    .Name = String.Empty
                    .Classification = String.Empty
                    .Star.Invalidate()
                    .Constellation = String.Empty
                    .Dimension = Single.NaN
                    .Magnitude = Single.NaN
                    .Description = String.Empty
                End With
                Return RetVal
            End Function
        End Structure

        Public ReadOnly Property FileName() As String
            Get
                Return MyFileName
            End Get
        End Property
        Private MyFileName As String = String.Empty

        Public Function GetEntry(ByVal NameToSearch As String) As sNGCEntry
            For Each Entry As sNGCEntry In Catalog
                If Entry.Name = NameToSearch Then Return Entry
            Next Entry
            Return sNGCEntry.Empty
        End Function

        Public Sub New(ByVal FileToLoad As String)

            Catalog = New List(Of sNGCEntry)
            Dim ErrorCount As Integer = 0

            If System.IO.File.Exists(FileToLoad) = False Then Exit Sub
            MyFileName = FileToLoad

            For Each BlockContent As String In System.IO.File.ReadAllLines(FileName, System.Text.Encoding.ASCII)
                Try
                    Dim NewEntry As New sNGCEntry
                    NewEntry.Name = BlockContent.Substring(0, 5).Trim.Replace("I", "IC") : If NewEntry.Name.StartsWith("I") = False Then NewEntry.Name = "NGC" & NewEntry.Name
                    NewEntry.Classification = TranslateClassification(BlockContent.Substring(6, 3))
                    Dim RA As Double = DegFromHMS(BlockContent.Substring(10, 2), BlockContent.Substring(13, 4))
                    Dim DE As Double = (CInt(BlockContent.Substring(20, 2)) + (CInt(BlockContent.Substring(23, 2)) / 60)) * CDbl(IIf(BlockContent.Substring(19, 1) = "-", -1, 1))
                    NewEntry.Star = New sGeneric(RA, DE)
                    NewEntry.Constellation = BlockContent.Substring(29, 3)
                    NewEntry.Dimension = CSng(Val(BlockContent.Substring(33, 5)))
                    NewEntry.Magnitude = CSng(Val(BlockContent.Substring(40, 4)))
                    NewEntry.Description = BlockContent.Substring(46).Replace(", ", ",").Trim
                    Catalog.Add(NewEntry)
                Catch ex As Exception
                    ErrorCount += 1
                End Try
            Next BlockContent

        End Sub

        Private Function TranslateClassification(ByVal Code As String) As String
            Select Case Code.Trim
                Case "Gx" : Return "Galaxy"
                Case "OC" : Return "Open star cluster"
                Case "Gb" : Return "Globular star cluster, usually in the Milky Way Galaxy"
                Case "Nb" : Return "Bright emission or reflection nebula"
                Case "Pl" : Return "Planetary nebula"
                Case "C+N" : Return "Cluster associated with nebulosity"
                Case "Ast" : Return "Asterism or group of a few stars"
                Case "Kt" : Return "Knot or nebulous region in an external galaxy"
                Case "***" : Return "Triple star"
                Case "D*" : Return "Double star"
                Case "*" : Return "Single star"
                Case "?" : Return "Uncertain type or may not exist"
                Case "-" : Return "Object called nonexistent in the RNGC (Sulentic and Tifft 1973)"
                Case "PD" : Return "Photographic plate defect"
                Case Else : Return "Unidentified at the place given, or type unknown"
            End Select
        End Function

    End Class

    '''<summary>A list of all available constellations (88 in sum).</summary>
    Public Structure sConstellation

        '''<summary>Name according to IAU.</summary>
        Public IAU As String
        '''<summary>Latin name.</summary>
        Public Latin As String
        '''<summary>Genitive name.</summary>
        Public Genitive As String
        '''<summary>English name.</summary>
        Public English As String
        '''<summary>Covered area [square-degree].</summary>
        Public Area As Integer
        '''<summary>Hemisphere.</summary>
        Public Hemisphere As String
        '''<summary>Main star name (only if special name).</summary>
        Public AlphaStar As String
        '''<summary>German name.</summary>
        Public German As String

        '''<summary>Constellation boundaries.</summary>
        Public Boundary As List(Of sGeneric)

        Public Sub New(ByVal NewIAU As String, ByVal NewLatin As String, ByVal NewGenitive As String, ByVal NewEnglish As String, ByVal NewArea As Integer, ByVal NewHemisphere As String, ByVal NewAlphaStar As String, ByVal NewGerman As String)
            Me.IAU = NewIAU
            Me.Latin = NewLatin
            Me.Genitive = NewGenitive
            Me.English = NewEnglish
            Me.Area = NewArea
            Me.Hemisphere = NewHemisphere
            Me.AlphaStar = NewAlphaStar
            Me.German = NewGerman
            Me.Boundary = New List(Of sGeneric)
        End Sub

        Public Shared Function Init(ByVal NewIAU As String, ByVal NewLatin As String, ByVal NewGenitive As String, ByVal NewEnglish As String, ByVal NewArea As Integer, ByVal NewHemisphere As String, ByVal NewAlphaStar As String) As sConstellation
            Return Init(NewIAU, NewLatin, NewGenitive, NewEnglish, NewArea, NewHemisphere, NewAlphaStar, String.Empty)
        End Function

        Public Shared Function Init(ByVal NewIAU As String, ByVal NewLatin As String, ByVal NewGenitive As String, ByVal NewEnglish As String, ByVal NewArea As Integer, ByVal NewHemisphere As String, ByVal NewAlphaStar As String, ByVal NewGerman As String) As sConstellation

            Dim RetVal As sConstellation
            RetVal.IAU = NewIAU
            RetVal.Latin = NewLatin
            RetVal.Genitive = NewGenitive
            RetVal.English = NewEnglish
            RetVal.Area = NewArea
            RetVal.Hemisphere = NewHemisphere
            RetVal.AlphaStar = NewAlphaStar
            RetVal.German = NewGerman
            RetVal.Boundary = New List(Of sGeneric)

            Return RetVal

        End Function

    End Structure

    '''<summary>Information related to constellations.</summary>
    Public Class cConstellation

        '''<summary>Shared catalog containing all constellations (hard-coded).</summary>
        Public Shared Catalog As New List(Of sConstellation)(New sConstellation() {
    New sConstellation("And", "Andromeda", "Andromedae", "Andromeda", 722, "NH", "Alpheratz", "Andromeda"),
     New sConstellation("Ant", "Antlia", "Antliae", "Air Pump", 239, "SH", "", "Luftpumpe"),
    New sConstellation("Aps", "Apus", "Apodis", "Bird of Paradise", 206, "SH", "", "Paradiesvogel"),
    New sConstellation("Aqr", "Aquarius", "Aquarii", "Water Carrier", 980, "SH", "Sadalmelik", "Adler"),
    New sConstellation("Aql", "Aquila", "Aquilae", "Eagle", 652, "NH-SH", "Altair", "Wassermann"),
    New sConstellation("Ara", "Ara", "Arae", "Altar", 237, "SH", "", "Altar"),
    New sConstellation("Ari", "Aries", "Arietis", "Ram", 441, "NH", "Hamal", "Widder"),
    New sConstellation("Aur", "Auriga", "Aurigae", "Charioteer", 657, "NH", "Capella", "Fuhrmann"),
    New sConstellation("Boo", "Bootes", "Bootis", "Herdsman", 907, "NH", "Arcturus", "Bärenhüter / Bootes"),
    New sConstellation("Cae", "Caelum", "Caeli", "Chisel", 125, "SH", "", "Grabstichel"),
    New sConstellation("Cam", "Camelopardalis", "Camelopardalis", "Giraffe", 757, "NH", "", "Giraffe"),
    New sConstellation("Cnc", "Cancer", "Cancri", "Crab", 506, "NH", "Acubens", "Krebs"),
    New sConstellation("CVn", "Canes Venatici", "Canun Venaticorum", "Hunting Dogs", 465, "NH", "Cor Caroli", "Jagdhunde"),
    New sConstellation("CMa", "Canis Major", "Canis Majoris", "Big Dog", 380, "SH", "Sirius", "Großer Hund"),
    New sConstellation("CMi", "Canis Minor", "Canis Minoris", "Little Dog", 183, "NH", "Procyon", "Kleiner Hund"),
    New sConstellation("Cap", "Capricornus", "Capricorni", "Goat ( Capricorn )", 414, "SH", "Algedi", "Steinbock"),
    New sConstellation("Car", "Carina", "Carinae", "Keel", 494, "SH", "Canopus", "Kiel des Schiffs"),
    New sConstellation("Cas", "Cassiopeia", "Cassiopeiae", "Cassiopeia", 598, "NH", "Schedar", "Kassiopeia"),
    New sConstellation("Cen", "Centaurus", "Centauri", "Centaur", 1060, "SH", "Rigil Kentaurus", "Zentaur"),
    New sConstellation("Cep", "Cepheus", "Cephei", "Cepheus", 588, "SH", "Alderamin", "Kepheus"),
    New sConstellation("Cet", "Cetus", "Ceti", "Whale", 1231, "SH", "Menkar", "Walfisch"),
    New sConstellation("Cha", "Chamaeleon", "Chamaleontis", "Chameleon", 132, "SH", "", "Chamäleon"),
    New sConstellation("Cir", "Circinus", "Circini", "Compasses", 93, "SH", "", "Zirkel"),
    New sConstellation("Col", "Columba", "Columbae", "Dove", 270, "SH", "Phact", "Taube"),
    New sConstellation("Com", "Coma Berenices", "Comae Berenices", "Berenice's Hair", 386, "NH", "Diadem", "Haar der Berenike"),
    New sConstellation("CrA", "Corona Australis", "Coronae Australis", "Southern Crown", 128, "SH", "", "Südliche Krone"),
    New sConstellation("CrB", "Corona Borealis", "Coronae Borealis", "Northern Crown", 179, "NH", "Alphecca", "Nördliche Krone"),
    New sConstellation("Crv", "Corvus", "Corvi", "Crow", 184, "SH", "Alchiba", "Rabe"),
    New sConstellation("Crt", "Crater", "Crateris", "Cup", 282, "SH", "Alkes", "Becher"),
    New sConstellation("Cru", "Crux", "Crucis", "Southern Cross", 68, "SH", "Acrux", "Kreuz des Südens"),
    New sConstellation("Cyg", "Cygnus", "Cygni", "Swan", 804, "NH", "Deneb", "Schwan"),
    New sConstellation("Del", "Delphinus", "Delphini", "Dolphin", 189, "NH", "Sualocin", "???"),
    New sConstellation("Dor", "Dorado", "Doradus", "Goldfish", 179, "SH", "", "???"),
    New sConstellation("Dra", "Draco", "Draconis", "Dragon", 1083, "NH", "Thuban", "???"),
    New sConstellation("Equ", "Equuleus", "Equulei", "Little Horse", 72, "NH", "Kitalpha", "???"),
    New sConstellation("Eri", "Eridanus", "Eridani", "River", 1138, "SH", "Achernar", "???"),
    New sConstellation("For", "Fornax", "Fornacis", "Furnace", 398, "SH", "", "???"),
    New sConstellation("Gem", "Gemini", "Geminorum", "Twins", 514, "NH", "Castor", "???"),
    New sConstellation("Gru", "Grus", "Gruis", "Crane", 366, "SH", "Al Na'ir", "???"),
    New sConstellation("Her", "Hercules", "Herculis", "Hercules", 1225, "NH", "Rasalgethi", "???"),
    New sConstellation("Hor", "Horologium", "Horologii", "Clock", 249, "SH", "", "???"),
    New sConstellation("Hya", "Hydra", "Hydrae", "Hydra ( Sea Serpent )", 1303, "SH", "Alphard", "???"),
    New sConstellation("Hyi", "Hydrus", "Hydri", "Water Serpen ( male )", 243, "SH", "", "???"),
    New sConstellation("Ind", "Indus", "Indi", "Indian", 294, "SH", "", "???"),
    New sConstellation("Lac", "Lacerta", "Lacertae", "Lizard", 201, "NH", "", "???"),
    New sConstellation("Leo", "Leo", "Leonis", "Lion", 947, "NH", "Regulus", "???"),
    New sConstellation("LMi", "Leo Minor", "Leonis Minoris", "Smaller Lion", 232, "NH", "", "???"),
    New sConstellation("Lep", "Lepus", "Leporis", "Hare", 290, "SH", "Arneb", "???"),
    New sConstellation("Lib", "Libra", "Librae", "Balance", 538, "SH", "Zubenelgenubi", "???"),
    New sConstellation("Lup", "Lupus", "Lupi", "Wolf", 334, "SH", "Men", "???"),
    New sConstellation("Lyn", "Lynx", "Lyncis", "Lynx", 545, "NH", "", "???"),
    New sConstellation("Lyr", "Lyra", "Lyrae", "Lyre", 286, "NH", "Vega", "???"),
    New sConstellation("Men", "Mensa", "Mensae", "Table", 153, "SH", "", "???"),
    New sConstellation("Mic", "Microscopium", "Microscopii", "Microscope", 210, "SH", "", "???"),
    New sConstellation("Mon", "Monoceros", "Monocerotis", "Unicorn", 482, "SH", "", "???"),
    New sConstellation("Mus", "Musca", "Muscae", "Fly", 138, "SH", "", "???"),
    New sConstellation("Nor", "Norma", "Normae", "Square", 165, "SH", "", "???"),
    New sConstellation("Oct", "Octans", "Octantis", "Octant", 291, "SH", "", "???"),
    New sConstellation("Oph", "Ophiuchus", "Ophiuchi", "Serpent Holder", 948, "NH-SH", "Rasalhague", "???"),
    New sConstellation("Ori", "Orion", "Orionis", "Orion", 594, "NH-SH", "Betelgeuse", "???"),
    New sConstellation("Pav", "Pavo", "Pavonis", "Peacock", 378, "SH", "Peacock", "???"),
    New sConstellation("Peg", "Pegasus", "Pegasi", "Winged Horse", 1121, "NH", "Markab", "???"),
    New sConstellation("Per", "Perseus", "Persei", "Perseus", 615, "NH", "Mirfak", "???"),
    New sConstellation("Phe", "Phoenix", "Phoenicis", "Phoenix", 469, "SH", "Ankaa", "???"),
    New sConstellation("Pic", "Pictor", "Pictoris", "Easel", 247, "SH", "", "???"),
    New sConstellation("Psc", "Pisces", "Piscium", "Fishes", 889, "NH", "Alrischa", "???"),
    New sConstellation("PsA", "Piscis Austrinus", "Piscis Austrini", "Southern Fish", 245, "SH", "Fomalhaut", "???"),
    New sConstellation("Pup", "Puppis", "Puppis", "Stern", 673, "SH", "", "???"),
    New sConstellation("Pyx", "Pyxis", "Pyxidis", "Compass", 221, "SH", "", "???"),
    New sConstellation("Ret", "Reticulum", "Reticuli", "Reticle", 114, "SH", "", "???"),
    New sConstellation("Sge", "Sagitta", "Sagittae", "Arrow", 80, "NH", "", "???"),
    New sConstellation("Sgr", "Sagittarius", "Sagittarii", "Archer", 867, "SH", "Rukbat", "???"),
    New sConstellation("Sco", "Scorpius", "Scorpii", "Scorpion", 497, "SH", "Antares", "???"),
    New sConstellation("Scl", "Sculptor", "Sculptoris", "Sculptor", 475, "SH", "", "???"),
    New sConstellation("Sct", "Scutum", "Scuti", "Shield", 109, "SH", "", "???"),
    New sConstellation("Ser", "Serpens", "Serpentis", "Serpent", 637, "NH-SH", "Unuck al Hai", "???"),
    New sConstellation("Sex", "Sextans", "Sextantis", "Sextant", 314, "SH", "", "???"),
    New sConstellation("Tau", "Taurus", "Tauri", "Bull", 797, "NH", "Aldebaran", "???"),
    New sConstellation("Tel", "Telescopium", "Telescopii", "Telescope", 252, "SH", "", "???"),
    New sConstellation("Tri", "Triangulum", "Trianguli", "Triangle", 132, "NH", "Ras al Mothallah", "???"),
    New sConstellation("TrA", "Triangulum Australe", "Trianguli Australis", "Southern Triangle", 110, "SH", "Atria", "???"),
    New sConstellation("Tuc", "Tucana", "Tucanae", "Toucan", 295, "SH", "", "???"),
    New sConstellation("UMa", "Ursa Major", "Ursae Majoris", "Great Bear", 1280, "NH", "Dubhe", "???"),
    New sConstellation("UMi", "Ursa Minor", "Ursae Minoris", "Little Bear", 256, "NH", "Polaris", "???"),
    New sConstellation("Vel", "Vela", "Velorum", "Sails", 500, "SH", "", "???"),
    New sConstellation("Vir", "Virgo", "Virginis", "Virgin", 1294, "NH-SH", "Spica", "???"),
    New sConstellation("Vol", "Volans", "Volantis", "Flying Fish", 141, "SH", "", "???"),
    New sConstellation("Vul", "Vulpecula", "Vulpeculae", "Fox", 268, "NH", "", "???")
  })

        Public Function Verbose(ByVal Abbreviation As String) As String
            Abbreviation = Abbreviation.ToUpper
            For Each Entry As sConstellation In Catalog
                If Entry.IAU.ToUpper = Abbreviation Then Return Entry.Latin
            Next Entry
            Return Abbreviation
        End Function

        Public Function InjectBoundaries(ByVal BoundaryFile As String) As Boolean

            If System.IO.File.Exists(BoundaryFile) = False Then Return False

            For Each Line As String In System.IO.File.ReadAllLines(BoundaryFile)

                Dim RAhr As Double = DegFromHMS(Line.Substring(0, 10))
                Dim DEdeg As Double = Val(Line.Substring(11, 11).Replace("+", String.Empty))
                Dim cst As String = Line.Substring(23, 4).Trim
                Dim type As String = Line.Substring(28, 1)

                For Each Entry As sConstellation In Catalog
                    If Entry.IAU.ToUpper = cst Then
                        Entry.Boundary.Add(New sGeneric(RAhr, DEdeg))
                    End If
                Next Entry

            Next Line

            DumpBoundaries(0)

            Return True

        End Function

        Private Sub DumpBoundaries(ByVal Idx As Integer)
            Dim DumpLine As New List(Of String)
            For Each Entry As sGeneric In Catalog.Item(Idx).Boundary
                DumpLine.Add(Entry.RightAscension.ToString.Trim & ";" & Entry.Declination.ToString.Trim)
            Next Entry
            System.IO.File.WriteAllLines("D:\Privat\Astro_TEMP\" & Catalog.Item(Idx).IAU & ".csv", DumpLine.ToArray)
        End Sub

        Public Sub New()
            'Translations from here: http://de.wikipedia.org/wiki/Liste_der_Sternbilder_in_verschiedenen_Sprachen
            Catalog = New List(Of sConstellation)
        End Sub

    End Class

    '=================================================================================================================
    ' HELPER FUNCTIONS
    '=================================================================================================================

    '''<summary>Convert the given hours (decimals allowed) to degree.</summary>
    Public Shared Function DegFromHMS(ByVal Hours As String) As Double
        Return Val(Hours.Replace(",", ".")) * 15
    End Function

    Public Shared Function DegFromHMS(ByVal Hours As String, ByVal Minutes As String) As Double
        Return (CInt(Hours) + (Val(Minutes) / 60)) * 15
    End Function

    Public Shared Function DegFromHMS(ByVal Hours As Integer, ByVal Minutes As Integer, ByVal Seconds As Double) As Double
        Return 15 * (Hours + (Minutes / 60) + (Seconds / 3600))
    End Function

    Public Shared Function DegFromDMS(ByVal Degree As Integer, ByVal Minutes As Integer, ByVal Seconds As Double) As Double
        Return Degree + (Minutes / 60) + (Seconds / 3600)
    End Function

    '''<summary>Convert the given greek leter description to a unicode character.</summary>
    Public Shared Function ToGreek(ByVal Text As String) As String

        Text = Text.Replace("alf", "α")
        Text = Text.Replace("bet", "β")
        Text = Text.Replace("gam", "γ")
        Text = Text.Replace("kap", "κ")
        Text = Text.Replace("eps", "ε")
        Text = Text.Replace("the", "θ")
        Text = Text.Replace("chi", "χ")
        Text = Text.Replace("sig", "σ")
        Text = Text.Replace("iot", "ι")
        Text = Text.Replace("zet", "ζ")
        Text = Text.Replace("rho", "ρ")
        Text = Text.Replace("lam", "λ")
        Text = Text.Replace("omi", "ο")
        Text = Text.Replace("eta", "η")
        Text = Text.Replace("phi", "φ")
        Text = Text.Replace("ome", "ω")
        Text = Text.Replace("tau", "τ")
        Text = Text.Replace("ups", "ϒ")
        Text = Text.Replace("psi", "ψ")
        Text = Text.Replace("pi", "π")
        Text = Text.Replace("del", "δ")

        Text = Text.Replace("01", "₁")
        Text = Text.Replace("02", "₂")
        Text = Text.Replace("03", "₃")
        Text = Text.Replace("04", "₄")
        Text = Text.Replace("05", "₅")
        Text = Text.Replace("06", "₆")
        Text = Text.Replace("07", "₇")
        Text = Text.Replace("08", "₈")
        Text = Text.Replace("09", "₉")

        Return Text

    End Function

    '''<summary>Get the constallation line for the requested constellation.</summary>
    '''<param name="IAUName">IAU name.</param>
    '''<returns>List of HD numbers - -1 indicated to insert a "break" in the line.</returns>
    Public Function GetConstLine(ByVal IAUName As String) As List(Of Integer)

        Dim RetVal As New List(Of Integer)

        Select Case IAUName.ToUpper

            Case "UMa".ToUpper
                'Ursa Major 
                RetVal.Add(120315)
                RetVal.Add(116656)
                RetVal.Add(112185)
                RetVal.Add(106591)
                RetVal.Add(95689)
                RetVal.Add(95418)
                RetVal.Add(103287)
                RetVal.Add(106591)

            Case Else

                'Do nothing...

        End Select

        Return RetVal

    End Function

End Class