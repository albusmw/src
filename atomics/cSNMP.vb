Option Explicit On
Option Strict On

Public Class cSNMP

  Public Property HostIP() As String
    Get
      Return MyHostIP
    End Get
    Set(value As String)
      MyHostIP = value
    End Set
  End Property
  Private MyHostIP As String = String.Empty

  Public Property GetCommunity() As String
    Get
      Return MyGetCommunity
    End Get
    Set(value As String)
      MyGetCommunity = value
    End Set
  End Property
  Private MyGetCommunity As String = "public"

  Public Property SetCommunity() As String
    Get
      Return MySetCommunity
    End Get
    Set(value As String)
      MySetCommunity = value
    End Set
  End Property
  Private MySetCommunity As String = "public"

  Public Property SnmpVersion() As SnmpSharpNet.SnmpVersion
    Get
      Return MySnmpVersion
    End Get
    Set(value As SnmpSharpNet.SnmpVersion)
      MySnmpVersion = value
    End Set
  End Property
  Private MySnmpVersion As SnmpSharpNet.SnmpVersion = SnmpSharpNet.SnmpVersion.Ver1

    Public Function GetCounter(ByVal OID As String) As Integer

        'Not implemented ...
        Return -1

    End Function

    Public Function GetInteger(ByVal OID As String) As Integer

        Dim RetVal As Integer = -1
        Dim snmp As SnmpSharpNet.SimpleSnmp = New SnmpSharpNet.SimpleSnmp(HostIP, GetCommunity)

    If snmp.Valid = True Then
      Dim requestOid() As String = New String() {OID}
      Dim Result As Dictionary(Of SnmpSharpNet.Oid, SnmpSharpNet.AsnType) = snmp.Get(SnmpVersion, requestOid)
      If Result IsNot Nothing Then
        Dim kvp As KeyValuePair(Of SnmpSharpNet.Oid, SnmpSharpNet.AsnType)
        For Each kvp In Result
          Try
                        RetVal = CInt(kvp.Value.ToString)
                    Catch ex As Exception
                        RetVal = -1
                    End Try
        Next kvp
      End If
    End If

    snmp = Nothing
        Return RetVal

    End Function

  Public Function SetInteger(ByVal OID As String, ByVal RequestedValue As Integer) As Integer

    Dim RetVal As Integer = -1

    'Prepare target
    Dim target = New SnmpSharpNet.UdpTarget(System.Net.IPAddress.Parse(HostIP))

    'Create a SET PDU
    Dim pdu = New SnmpSharpNet.Pdu(SnmpSharpNet.PduType.Set)

    'Set value to a integer
    pdu.VbList.Add(New SnmpSharpNet.Oid(OID), New SnmpSharpNet.Integer32(RequestedValue))

    'Set Agent security parameters
    Dim aparam = New SnmpSharpNet.AgentParameters(SnmpVersion, New SnmpSharpNet.OctetString(SetCommunity))

    'Response packet
    Dim response As SnmpSharpNet.SnmpPacket = Nothing

    Try
      'Send request and wait for response
      response = target.Request(pdu, aparam)
    Catch ex As Exception
      'Do nothing ...
    End Try

    'Make sure we received a response
    If IsNothing(response) = False Then
      'Check if we received an SNMP error from the agent
      If response.Pdu.ErrorStatus = 0 Then
        'Everything is ok. Agent will return the new value for the OID we changed
        Try
          RetVal = CInt(response.Pdu(0).Value.ToString)
        Catch ex As Exception
          'Do nothing ...
        End Try
      End If
    End If

    target.Close()

    Return RetVal

  End Function


End Class
