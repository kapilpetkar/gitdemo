'*******
'** TC001 : Test class
'**         This is just a test class to check http session state.
'*******
Option Strict On
Option Explicit On

Public Class CurrentSession

    Private Shared l001_session As HttpSessionState
    Public Shared Property session() As HttpSessionState
        Get
            Return l001_session
        End Get
        Set(ByVal value As HttpSessionState)
            l001_session = value
        End Set
    End Property

End Class
