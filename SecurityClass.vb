Imports System.IO
Imports System.Management
Imports System.Threading.Tasks
Imports System.Net

Namespace MyNamespace
    Public Class SecurityClass
        Public Shared Function GetCpuId() As String
            Dim cpuId As String = String.Empty
            Try
                Dim searcher As New ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor")
                For Each obj As ManagementObject In searcher.Get()
                    cpuId = obj("ProcessorId").ToString()
                    Exit For ' Retrieve the first processor's ID
                Next
            Catch
                cpuId = "Error retrieving CPU ID"
            End Try
            Return cpuId
        End Function
        Public Shared Function IsAllowed() As Boolean
            Dim cpuId As String = GetCpuId()
            Dim title As String = String.Empty
            Dim URL As String = "https://raw.githubusercontent.com/tomer7X/VerifyLicense/master/IDS.txt"
            ' Get HTML data
            Dim client As WebClient = New WebClient()
            Dim data As Stream = client.OpenRead(URL)
            Dim reader As StreamReader = New StreamReader(data)
            Dim str As String = ""
            str = reader.ReadLine()
            Do While str <> "END"
                If cpuId = str Then
                    Return True
                End If
                str = reader.ReadLine()
            Loop
            Return False
        End Function
    End Class
End Namespace