Option Explicit On
Option Strict On

'''<summary>Class to control downloading content from the internet.</summary>
Public Class cDownloader

    Public Shared Sub UseNoProxy()
        MyProxyURL = String.Empty
        MyProxyPort = 0
    End Sub

    Public Shared Sub ConfigureProxy(ByVal ProxyURL As String, ByVal ProxyPort As Integer, ByVal UserName As String, ByVal Password As String)
        MyProxyURL = ProxyURL
        MyProxyPort = ProxyPort
        MyUserName = UserName
        MyPassword = Password
    End Sub
    Private Shared MyProxyURL As String = String.Empty
    Private Shared MyProxyPort As Integer = 0
    Private Shared MyUserName As String = String.Empty
    Private Shared MyPassword As String = String.Empty

    '''<summary>Web client to be used.</summary>
    Private Downloader As System.Net.WebClient

    '''<summary>Init the web client for download.</summary>
    Public Sub InitWebClient()
        If IsNothing(Downloader) Then Downloader = New System.Net.WebClient()
        If String.IsNullOrEmpty(MyProxyURL) = False Then ConfigProxy(Downloader)
        Downloader.Encoding = System.Text.Encoding.UTF8
    End Sub

    'Configure the connection to work with the specified proxy
    Public Sub ConfigProxy(ByRef Connection As System.Net.WebClient)
        Dim pr As New System.Net.WebProxy(MyProxyURL, MyProxyPort)
        Connection.Proxy = pr
        If String.IsNullOrEmpty(MyUserName) = False Then
            Dim cr As New System.Net.NetworkCredential(MyUserName, MyPassword)
            Connection.Proxy.Credentials = cr
        End If
    End Sub

    Public Function DownloadString(ByVal URL As String) As String
        If IsNothing(Downloader) Then InitWebClient()
        Return Downloader.DownloadString(URL)
    End Function

    Public Function DownloadFile(ByVal URL As String, ByVal FileName As String) As Boolean
        Try
            Downloader.DownloadFile(URL, FileName)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class

