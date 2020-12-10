Option Explicit On
Option Strict On

'Dictionary to handle status results
Public Class cDictEx

    '''<summary>Object to store values in the database with a date and time fingerprint.</summary>
    Private Class cValueType
        Public Property Value As Object
        Public Property LastWrite As DateTime
        Public Sub New(ByVal NewValue As Object)
            Me.Value = NewValue
            If IsNothing(NewValue) = False Then
                Me.LastWrite = Now
            Else
                Me.LastWrite = DateTime.MinValue
            End If
        End Sub
    End Class

    Private MyDict As New Dictionary(Of DB.eKey, cValueType)

    '''<summary>Set (or add if not exists) a new key value pair.</summary>
    Public Sub [Set](ByVal Key As DB.eKey, ByVal Value As Object)
        If MyDict.ContainsKey(Key) = False Then
            'Key did not exist -> add
            MyDict.Add(Key, New cValueType(Value))
        Else
            'Key did exist -> add
            MyDict(Key) = New cValueType(Value)
        End If
    End Sub

    '''<summary>Get the value for a certain key value.</summary>
    Default Public ReadOnly Property [Get](ByVal Key As DB.eKey) As Object
        Get
            If MyDict.ContainsKey(Key) Then
                Return MyDict(Key).Value
            Else
                Return Nothing
            End If
        End Get
    End Property

    '''<summary>Set (or add if not exists) a new key value pair.</summary>
    Public Sub [Clear](ByVal Key As DB.eKey)
        If MyDict.ContainsKey(Key) = True Then
            'Key did exist -> add
            MyDict.Remove(Key)
        End If
    End Sub

    '''<summary>Get the age (=different from last update to now) of the data in milliseconds.</summary>
    Public ReadOnly Property GetDataAge(ByVal Key As DB.eKey) As Long
        Get
            If MyDict.ContainsKey(Key) Then                                             'value is present
                If MyDict(Key).LastWrite <> DateTime.MinValue Then                      'value was written at any time
                    Return CLng(Now.Subtract(MyDict(Key).LastWrite).TotalMilliseconds)
                End If
            End If
            Return -1
        End Get
    End Property

    '''<summary>Get a database key in a specific type (defined by the default type).</summary>
    '''<param name="Key">Key to load.</param>
    '''<param name="[Default]">Default value (also defines the return type).</param>
    '''<returns>Value of the key or the default value in case of key-not-found or wrong type.</returns>
    Public Function GetTyped(ByVal Key As DB.eKey, ByVal [Default] As Boolean) As Boolean
        If MyDict.ContainsKey(Key) = False Then Return [Default]
        Try
            Return CBool(MyDict(Key).Value)
        Catch ex As Exception
            Return [Default]
        End Try
    End Function

    '''<summary>Get a database key in a specific type (defined by the default type).</summary>
    '''<param name="Key">Key to load.</param>
    '''<param name="[Default]">Default value (also defines the return type).</param>
    '''<returns>Value of the key or the default value in case of key-not-found or wrong type.</returns>
    Public Function GetTyped(ByVal Key As DB.eKey, ByVal [Default] As String) As String
        If MyDict.ContainsKey(Key) = False Then Return [Default]
        Try
            Return CStr(MyDict(Key).Value)
        Catch ex As Exception
            Return [Default]
        End Try
    End Function

    '''<summary>Get a database key in a specific type (defined by the default type).</summary>
    '''<param name="Key">Key to load.</param>
    '''<param name="[Default]">Default value (also defines the return type).</param>
    '''<returns>Value of the key or the default value in case of key-not-found or wrong type.</returns>
    Public Function GetTyped(ByVal KeyAsString As String, ByVal [Default] As String) As String
        Dim Key As DB.eKey
        If DB.eKey.TryParse(KeyAsString, Key) = True Then
            If MyDict.ContainsKey(Key) = False Then Return [Default]
            Try
                Return CStr(MyDict(Key).Value)
            Catch ex As Exception
                Return [Default]
            End Try
        Else
            Return [Default]
        End If
    End Function

End Class