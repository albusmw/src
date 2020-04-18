Option Explicit On
Option Strict On

Public Class cSNMP

    Public Property HostIP() As String = String.Empty
    Public Property GetCommunity() As String = "private"
    Public Property SetCommunity() As String = "private"
    Public Property SnmpVersion() As SnmpSharpNet.SnmpVersion = SnmpSharpNet.SnmpVersion.Ver1

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

    Public Function GetDouble(ByVal OID As String) As Double

        Dim RetVal As Double = Double.NaN
        Dim snmp As SnmpSharpNet.SimpleSnmp = New SnmpSharpNet.SimpleSnmp(HostIP, GetCommunity)

        If snmp.Valid = True Then
            Dim requestOid() As String = New String() {OID}
            Dim Result As Dictionary(Of SnmpSharpNet.Oid, SnmpSharpNet.AsnType) = snmp.Get(SnmpVersion, requestOid)
            If Result IsNot Nothing Then
                Dim kvp As KeyValuePair(Of SnmpSharpNet.Oid, SnmpSharpNet.AsnType)
                For Each kvp In Result
                    Try
                        RetVal = CDbl(kvp.Value.ToString)
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

        Dim target = New SnmpSharpNet.UdpTarget(System.Net.IPAddress.Parse(HostIP))                             'Prepare target
        Dim pdu = New SnmpSharpNet.Pdu(SnmpSharpNet.PduType.Set)                                                'Create a SET PDU
        Dim aparam = New SnmpSharpNet.AgentParameters(SnmpVersion, New SnmpSharpNet.OctetString(SetCommunity))  'Set Agent security parameters

        'Response packet
        pdu.VbList.Add(New SnmpSharpNet.Oid(OID), New SnmpSharpNet.Integer32(RequestedValue))                   'Set Agent security parameters'Set value to a integer
        Dim response As SnmpSharpNet.SnmpPacket = Nothing
        Dim LastError As Exception = Nothing

        Try
            'Send request and wait for response
            response = target.Request(pdu, aparam)
        Catch ex As Exception
            LastError = ex
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
