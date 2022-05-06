Imports MyStudenApp.WebReference

Public Class SFDCHelper
    Public m_Sforce As SforceService
    Public m_User As String
    Public m_Password As String
    Public m_SecurityToken As String
    Public m_ProxyHost As String
    Public m_ProxyPort As String
    Public m_ProxyUser As String
    Public m_ProxyPass As String
    Public m_Url As String

    Private Const UPD_MAX_CNT As Integer = 200
    Private Const m_RetryCnt As Integer = 5
    Private Const m_SleepTime As Integer = 1000
    Private Const m_SleepExTime As Integer = 2000
    Private Const m_LoginTimeout As Integer = 10000000

    Public Sub New(ByVal user As String,
                   ByVal password As String,
                   ByVal securityToken As String,
                   ByVal proxyHost As String,
                   ByVal proxyPort As String,
                   ByVal proxyUser As String,
                   ByVal proxyPass As String
                   )
        Me.m_User = user
        Me.m_Password = password
        Me.m_SecurityToken = securityToken
        Me.m_ProxyHost = proxyHost
        Me.m_ProxyPort = proxyPort
        Me.m_ProxyUser = proxyUser
        Me.m_ProxyPass = proxyPass
        Me.m_Url = m_Sforce.Url
    End Sub

    Public Function login() As String
        Dim isLoginTimeOut As Boolean = False
        Try
            Dim timeStamp As Date = m_Sforce.getServerTimestamp.timestamp

        Catch ex As Exception
            isLoginTimeOut = True
        End Try

        If Not isLoginTimeOut Then
            Return True
        End If

        Try
            Me.m_Sforce.Timeout = m_LoginTimeout
            If Not String.IsNullOrEmpty(m_ProxyHost) And Not String.IsNullOrEmpty(m_ProxyPort) Then
                Me.m_Sforce.Proxy = New Net.WebProxy(m_ProxyHost, CInt(m_ProxyPort))
            End If
            If Not String.IsNullOrEmpty(m_ProxyUser) And Not String.IsNullOrEmpty(m_ProxyPass) Then
                Me.m_Sforce.Proxy.Credentials = New Net.NetworkCredential(m_ProxyUser, m_ProxyPass)
            End If
            Dim loginResult As LoginResult = Me.m_Sforce.login(Me.m_User, m_Password + m_SecurityToken)
            Me.m_Sforce.SessionHeaderValue = New SessionHeader
            Me.m_Sforce.SessionHeaderValue.sessionId = loginResult.sessionId
            Me.m_Sforce.Url = loginResult.serverUrl
            MainController.strExecutionUser = loginResult.userId
            Return ""
        Catch ex As Exception
            Return ex.Message()
        End Try
    End Function
End Class
