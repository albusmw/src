Option Explicit On
Option Strict On

Namespace AstroCalc.NET

  Namespace Databases

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
      Public ReadOnly Property LocalFolder() As String
        Get
          Return MyLocalFolder
        End Get
      End Property
      Private MyLocalFolder As String = String.Empty

      '''<summary>Downloader for content.</summary>
      Private MyDownloader As New AstroCalc.NET.Common.cDownloader

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

        MyLocalFolder = CatalogFolder
        If System.IO.File.Exists(MyLocalFolder & "\" & ExtractedFile) = False Then Exit Sub

        'TODO: This is not the finalized code ...

        For Each BlockContent As String In System.IO.File.ReadAllLines(MyLocalFolder & "\" & ExtractedFile, System.Text.Encoding.ASCII)
          'Add some spaces in the back to avoid problems during parsing
          BlockContent &= "                           "
          Try
            Dim NewEntry As New sHIPEntry
            NewEntry.HIP = CInt(BlockContent.Substring(8, 6).Trim)
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

  End Namespace

End Namespace