Option Explicit On
Option Strict On

Namespace AstroCalc.NET

  Namespace Databases

    Public Class cCrossIndexKostjuk

      ' Byte-by-byte Description of file: catalog.dat
      ' --------------------------------------------------------------------------------
      '    Bytes Format Units   Label   Explanations
      ' --------------------------------------------------------------------------------
      '    1-  6  I6    ---     HD      [1/257937] Henry Draper Catalog Number <III/135>
      '    8- 19  A12   ---     DM      Durchmusterung Identification from HD Catalog
      '                                   <III/135> (1)
      '   21- 25  I5    ---     GC      [1/33342]? Boss General Catalog (GC, <I/113>)
      '                                    number if one exists, otherwise blank
      '   27- 30  I4    ---     HR      [1/9110]? Harvard Revised Number=Bright Star
      '                                    Number <V/50> if one exists, otherwise blank
      '   32- 37  I6    ---     HIP     [1/120416]? Hipparcos Catalog <I/196> number
      '                                    if one exists, otherwise blank
      '   39- 40  I2    h       RAh     Right Ascension J2000 (hours) (2)
      '   41- 42  I2    min     RAm     Right Ascension J2000 (minutes) (2)
      '   43- 47  F5.2  s       RAs     Right Ascension J2000 (seconds) (2)
      '       49  A1    ---     DE-     Declination J2000 (sign)
      '   50- 51  I2    deg     DEd     Declination J2000 (degrees) (2)
      '   52- 53  I2    arcmin  DEm     Declination J2000 (minutes) (2)
      '   54- 57  F4.1  arcsec  DEs     Declination J2000 (seconds) (2)
      '   59- 63  F5.2  mag     Vmag    [-1.44/13.4]? Visual magnitude (2)
      '   65- 67  I3    ---     Fl      ? Flamsteed number (G1)
      '   69- 73  A5    ---     Bayer   Bayer designation (G1)
      '   75- 77  A3    ---     Cst     Constellation abbreviation (G1)
      ' --------------------------------------------------------------------------------

      Public Structure sCrossIndexKostjuk
        '''<summary>Generic information.</summary>
        Public Property Star As sGeneric
        '''<summary>Henry Draper Catalog Number.</summary>
        Public HD As Integer
        '''<summary>Harvard Revised Number=Bright Star Number.</summary>
        Public HR As Integer
        '''<summary>Hipparcos Catalog Number.</summary>
        Public HIP As Integer
        '''<summary>Flamsteed number.</summary>
        Public Fl As Integer
        '''<summary>Bayer designation.</summary>
        Public Bayer As String
        '''<summary>Constellation abbreviation.</summary>
        Public Cst As String
        Public Sub Clear()
          Me.HD = -1
          Me.HR = -1
          Me.HIP = -1
          Me.Fl = -1
          Me.Bayer = String.Empty
          Me.Cst = String.Empty
        End Sub
      End Structure

      Public Catalog As New List(Of sCrossIndexKostjuk)

      '''<summary>Root URL where to load the content from.</summary>
      Public Property RootURL() As String
        Get
          Return MyRootURL
        End Get
        Set(value As String)
          MyRootURL = value
        End Set
      End Property
      Private MyRootURL As String = "ftp://cdsarc.u-strasbg.fr/pub/cats/IV/27A/catalog.dat"

      '''<summary>Downloader for content.</summary>
      Private MyDownloader As New AstroCalc.NET.Common.cDownloader

      Public Event Currently(ByVal Info As String)

      '''<summary>Original file name (as present on server).</summary>
      Public ReadOnly Property OrigFileName() As String
        Get
          Return "catalog.dat"
        End Get
      End Property

      '''<summary>Extracted file name (as present on local disc after extract).</summary>
      Public ReadOnly Property ExtractedFile() As String
        Get
          Return "catalog.dat"
        End Get
      End Property

      '''<summary>Local folder where to find the catalog file.</summary>
      Public ReadOnly Property LocalFolder() As String
        Get
          Return MyLocalFolder
        End Get
      End Property
      Private MyLocalFolder As String = String.Empty

      Public Sub Download()
        MyDownloader.InitWebClient()
        RaiseEvent Currently("Loading Kostjuk Cross Index data from VIZIER FTP ...")
        MyDownloader.DownloadFile(RootURL, LocalFolder & "\" & OrigFileName)
        RaiseEvent Currently("   DONE.")
      End Sub

      Public Sub New()
        Me.New(String.Empty)
      End Sub

      Public Sub New(ByVal CatalogFolder As String)

        Dim BayerList As New List(Of String)        'All Bayer descriptors
        Dim ErrorCount As Integer = 0

        Catalog.Clear()

        MyLocalFolder = CatalogFolder
        If System.IO.File.Exists(MyLocalFolder & "\" & ExtractedFile) = False Then Exit Sub

        For Each BlockContent As String In System.IO.File.ReadAllLines(MyLocalFolder & "\" & ExtractedFile, System.Text.Encoding.ASCII)
          'Add some spaces in the back to avoid problems during parsing
          BlockContent &= "                           "
          Try

            Dim NewEntry As New sCrossIndexKostjuk

            NewEntry.HD = CatParser.GetInt(BlockContent, 1, 6)
            NewEntry.HR = CatParser.GetInt(BlockContent, 27, 30)
            NewEntry.HIP = CatParser.GetInt(BlockContent, 32, 37)
            NewEntry.Fl = CatParser.GetInt(BlockContent, 65, 67)
            NewEntry.Bayer = CatParser.GetString(BlockContent, 69, 73)
            NewEntry.Cst = CatParser.GetString(BlockContent, 75, 77)

            Dim RA As Double = (CatParser.GetInt(BlockContent, 39, 40) + (CatParser.GetInt(BlockContent, 41, 42) / 60) + (CatParser.GetFloat(BlockContent, 43, 47) / 3600)) * 15
            Dim DE As Double = (CatParser.GetInt(BlockContent, 50, 51) + (CatParser.GetInt(BlockContent, 52, 53) / 60) + (CatParser.GetInt(BlockContent, 54, 57) / 3600)) * CDbl(IIf(CatParser.GetString(BlockContent, 49) = "-", -1, 1))

            NewEntry.Star = New sGeneric(RA, DE)

            Catalog.Add(NewEntry)
            If BayerList.Contains(cBayer.ToGreek(NewEntry.Bayer)) = False Then BayerList.Add(cBayer.ToGreek(NewEntry.Bayer))

            'If NewEntry.Bayer = "V636" Then MsgBox("...")

          Catch ex As Exception
            ErrorCount += 1
          End Try
        Next BlockContent

      End Sub

      Public Function GetViaHD(ByVal HD As Integer) As sGeneric
        For Each Entry As sCrossIndexKostjuk In Catalog
          If Entry.HD = HD Then Return Entry.Star
        Next Entry
        Return Nothing
      End Function

    End Class

  End Namespace

End Namespace