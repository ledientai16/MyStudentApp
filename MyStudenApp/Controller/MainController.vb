Imports Ksvc.Utility
Public Class MainController
    Public Shared strExecutionUser As String
    Public Shared Sub MainProccess()
        Dim dtToday As Date
        dtToday = Now()
        If sysSetting.getSystemSettings(settingInfo) = -1 Then
            LogFileHelper.writeExecLog("get SystemSetting error", String.Format(Constant.LogErrorFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        End If
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        LogFileHelper.writeExecLog("get SystemSetting successful", String.Format(Constant.LogExecuteFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))

        If sysSetting.getLoginInfo(loginInfo) = -1 Then
            LogFileHelper.writeExecLog("get login error", String.Format(Constant.LogErrorFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        End If
        LogFileHelper.writeExecLog("get login info success", String.Format(Constant.LogExecuteFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
        Console.WriteLine(loginInfo.Item(Constant.UserNameTag))

        sfdcService = New SFDCHelper(loginInfo.Item(Constant.UserNameTag),
                                    loginInfo.Item(Constant.PassWordTag),
                                    loginInfo.Item(Constant.SecurityTokenTag),
                                    loginInfo.Item(Constant.ProxyHostTag),
                                    loginInfo.Item(Constant.ProxyPortTag),
                                    loginInfo.Item(Constant.ProxyUserTag),
                                    loginInfo.Item(Constant.ProxyPassTag)
                                    )
        If Not String.IsNullOrEmpty(sfdcService.login()) Then
            LogFileHelper.writeExecLog("login fail", String.Format(Constant.LogErrorFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        Else
            LogFileHelper.writeExecLog("login success", String.Format(Constant.LogErrorFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
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
