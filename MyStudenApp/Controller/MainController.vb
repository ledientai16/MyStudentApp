Imports Ksvc.Utility
Public Class MainController
    Public Shared strExecutionUser As String
    Public Shared logErrorFileName As String
    Public Shared logExecuteFileName As String
    Public Shared Sub MainProccess()
        If Not sysSetting.getSystemSettings(settingInfo) Then
            LogFileHelper.writeExecLog("get SystemSetting error", logErrorFileName, True, settingInfo.Item(Constant.LogErrorFileName))
            Environment.Exit(1)
        End If
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        LogFileHelper.writeExecLog("get SystemSetting successful", logExecuteFileName, True, settingInfo.Item(Constant.LogExecuteFileName))
        If Not sysSetting.getLoginInfo(loginInfo) Then
            LogFileHelper.writeExecLog("get LoginInfo error", logExecuteFileName, True, settingInfo.Item(Constant.LogExecuteFileName))
            Environment.Exit(1)
        End If
        LogFileHelper.writeExecLog("Get login info success", logExecuteFileName, True, settingInfo(Constant.LoginInfo))
        sfdcService = New SFDCHelper(loginInfo.Item(Constant.UserNameTag),
                                    loginInfo.Item(Constant.PassWordTag),
                                    loginInfo.Item(Constant.SecurityTokenTag),
                                    loginInfo.Item(Constant.ProxyHostTag),
                                    loginInfo.Item(Constant.ProxyPortTag),
                                    loginInfo.Item(Constant.ProxyUserTag),
                                    loginInfo.Item(Constant.ProxyPassTag)
                                    )
        If Not String.IsNullOrEmpty(sfdcService.login()) Then
            LogFileHelper.writeExecLog("Login error", "logExecuteFileName", True, settingInfo.Item(Constant.LogExecuteFileName))
            Environment.Exit(1)
        Else
            upserths()
            copycsv()
            deletefile()

        End If
    End Sub

    Private Shared Sub deletefile()
        Throw New NotImplementedException()
    End Sub

    Private Shared Sub copycsv()
        Throw New NotImplementedException()
    End Sub

    Private Shared Sub upserths()
        Throw New NotImplementedException()
    End Sub
End Class
