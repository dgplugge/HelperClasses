' Author: Donald G Plugge
' Date: 12/9/08
' Authentication of credentials
Public Class UserAccess

    ' dgp rev 7/12/07 Prefix state 
    Public Enum AccessMethod
        [WinAuth] = 0
        Impersonate = 1
        VMS = 2
    End Enum

    ' dgp rev 7/12/07 Prefix state 
    Public Enum Server
        SPIFFY = 1
        '        "NT-EIB-10-6B16" = 0
    End Enum

    Private mWinDomain As String = "NIH"
    Private mDomain As String = "nci.nih.gov"

    ' dgp rev 12/9/08 Authentication password
    Private mUsername As String
    Public Property Username() As String
        Get
            Return mUsername
        End Get
        Set(ByVal value As String)
            mUsername = value
        End Set
    End Property

    ' dgp rev 12/9/08 Authentication password
    Private mPassword As String
    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property

End Class
