Imports System.IO
Imports Ksvc.Utility
Imports MyStudenApp.WebReference

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
        Dim csvFileList As List(Of String) = CsvFileHelper.getCsvFileList(Constant.RootDirectory, "*")

        For Each csvFile As String In csvFileList

            Dim csvRowErrList As Dictionary(Of String, String) = New Dictionary(Of String, String)
            Dim csvRowList As List(Of String()) = CsvFileHelper.getCsvRowList(csvFile, csvRowErrList)

            If Not csvRowList.Any Then
                Continue For
            End If

            Dim startTimeReadCSV As DateTime = DateTime.Now
            Dim fileName As String = Path.GetFileName(csvFile)
            Dim msgFileNameCSV = String.Format(Constant.CSVFileNameMsg, fileName)


            Dim csvRecordSuccessList As New List(Of String())
            Dim csvRecordErrorList As New List(Of String())
            Dim rowHeader As String() = csvRowList(0)

            Dim studentUpsertList As New List(Of HocSinh__c)

            For i As Integer = 1 To csvRowList.Count - 1
                Dim csvRow As String() = csvRowList(i)
                Dim objHocSinh As HocSinh__c = HocSinhDAO.createHocSinh(csvRow)
                studentUpsertList.Add(objHocSinh)
            Next

            If studentUpsertList.Any Then
                Dim studentUpsertResultList As List(Of UpsertResult) = HocSinhDAO.upsertHocSinh(studentUpsertList)
                For i As Integer = 0 To studentUpsertResultList.Count - 1
                    If studentUpsertResultList(i).success Then
                        csvRecordSuccessList.Add(csvRowList(i + 1))
                    Else
                        csvRecordErrorList.Add(csvRowList(i + 1))
                    End If
                Next
            End If

            ' create csv file success
            If csvRecordSuccessList.Any Then
                csvRecordSuccessList.Insert(0, rowHeader)
                Dim csvSuccessName As String = String.Format(Constant.CSVSuccessFileName, startTimeReadCSV.ToString(Constant.DateTimeFortmat), Constant.CsvFileExtension)
                CsvFileHelper.createFileCSV(settingInfo.Item(Constant.BackupCsvFoler), csvSuccessName, csvRecordSuccessList)
            End If

            ' create csv file error
            If csvRecordErrorList.Any Then
                csvRecordErrorList.Insert(0, rowHeader)
                Dim csvErrorName As String = String.Format(Constant.CSVErrorFileName, startTimeReadCSV.ToString(Constant.DateTimeFortmat), Constant.CsvFileExtension)
                CsvFileHelper.createFileCSV(settingInfo.Item(Constant.ErrorCsvFolder), csvErrorName, csvRecordErrorList)
            End If
        Next

    End Sub
End Class
