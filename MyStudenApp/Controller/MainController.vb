Imports System.IO
Imports Ksvc.Utility
Imports MyStudenApp.WebReference

Public Class MainController
    Public Shared strExecutionUser As String
    Public Shared Sub MainProccess()
        '' get and save settingInfo from SettingInfo.xml
        If sysSetting.getSystemSettings(settingInfo) = -1 Then
            LogFileHelper.writeExecLog("get SystemSetting error", String.Format(Constant.LogErrorFileName, Now()), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        End If

        ''set protocol
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        LogFileHelper.writeExecLog("get SystemSetting successful", String.Format(Constant.LogExecuteFileName, Date.Now), True, settingInfo.Item(Constant.LogFolderName))

        '' get and save loginInfo from loginInfo.xml
        If sysSetting.getLoginInfo(loginInfo) = -1 Then
            LogFileHelper.writeExecLog("get login error", String.Format(Constant.LogErrorFileName, Date.Now), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        End If
        LogFileHelper.writeExecLog("get login info success", String.Format(Constant.LogExecuteFileName, Date.Now), True, settingInfo.Item(Constant.LogFolderName))
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
            LogFileHelper.writeExecLog("login fail", String.Format(Constant.LogErrorFileName, Date.Now), True, settingInfo.Item(Constant.LogFolderName))
            Environment.Exit(1)
        Else
            LogFileHelper.writeExecLog("login success", String.Format(Constant.LogExecuteFileName, Date.Now), True, settingInfo.Item(Constant.LogFolderName))
            upserths()
            copycsv()
            deletefile()
        End If
    End Sub

    Private Shared Sub deletefile()
        Dim sourceDir = settingInfo.Item(Constant.InputFolder)

        Dim pathFileList As String() = Directory.GetFiles(sourceDir, "*.csv")
        For Each f As String In pathFileList
            File.Delete(f)
        Next
    End Sub

    Private Shared Sub copycsv()
        Dim sourceDir = settingInfo.Item(Constant.InputFolder)
        Dim pathFileList = Directory.GetFiles(sourceDir, "*.csv")
        For Each f As String In pathFileList
            Dim fName = Path.GetFileNameWithoutExtension(f)

            Dim backupFileName = String.Format("{0}_{1: yyyyMMddHHmmss}.csv", fName, Date.Now)
            Dim backupDir = Constant.RootDirectory & "backup"

            Dim dPath = Path.Combine(backupDir, backupFileName)

            File.Copy(f, dPath, True)
        Next

    End Sub

    Public Shared Sub upserths()
        Dim csvFileList As List(Of String) = CsvFileHelper.getCsvFileList(settingInfo.Item(Constant.InputFolder), "*")

        Console.WriteLine(csvFileList.Count)
        For Each csvFile As String In csvFileList

            Dim csvRowErrList As Dictionary(Of String, String) = New Dictionary(Of String, String)
            Dim csvRowList As List(Of String()) = CsvFileHelper.getCsvRowList(csvFile, csvRowErrList)
            Console.WriteLine(csvRowList.Count)

            If Not csvRowList.Any Then
                Continue For
            End If

            Dim startTimeReadCSV As DateTime = DateTime.Now
            Dim fileName As String = Path.GetFileName(csvFile)
            Dim msgFileNameCSV = String.Format(Constant.CSVFileNameMsg, fileName)

            Dim studentUpsertList As New List(Of HocSinh__c)
            Dim csvRecordSuccessList As New List(Of String())
            Dim csvRecordErrorList As New List(Of String())
            Dim rowHeader As String() = csvRowList(0)

            Dim upsertErrorList As New List(Of Error__c)

            For i As Integer = 1 To csvRowList.Count - 1
                Dim csvRow As String() = csvRowList(i)
                Dim objHocSinh As HocSinh__c = HocSinhDAO.CreateHocSinh(csvRow)
                studentUpsertList.Add(objHocSinh)
            Next

            If studentUpsertList.Any Then
                Dim studentUpsertResultList As List(Of UpsertResult) = HocSinhDAO.upsertHocSinh(studentUpsertList)
                For i As Integer = 0 To studentUpsertResultList.Count - 1
                    If studentUpsertResultList(i).success Then
                        csvRecordSuccessList.Add(csvRowList(i + 1))
                    Else
                        Dim objError As Error__c = ErrorDAO.CreateError(fileName, i + 1, studentUpsertResultList(i).errors(0).message)
                        Console.WriteLine(studentUpsertResultList(i).errors(0).message)
                        upsertErrorList.Add(objError)
                        csvRecordErrorList.Add(csvRowList(i + 1))
                    End If
                Next
            End If

            ' Insert list Error__c to salesforce 
            If upsertErrorList.Any Then
                ErrorDAO.InsertError(upsertErrorList)
            End If

            ' create csv file success
            If csvRecordSuccessList.Any Then
                csvRecordSuccessList.Insert(0, rowHeader)
                Dim csvSuccessName As String = String.Format(Constant.CSVSuccessFileName, startTimeReadCSV.ToString(Constant.DateTimeFortmat), Constant.CSV_FILE_EXTENSION)
                CsvFileHelper.createFileCSV(settingInfo.Item(Constant.BackupCsvFoler), csvSuccessName, csvRecordSuccessList)
            End If

            ' create csv file error
            If csvRecordErrorList.Any Then
                csvRecordErrorList.Insert(0, rowHeader)
                Dim csvErrorName As String = String.Format(Constant.CSVErrorFileName, startTimeReadCSV.ToString(Constant.DateTimeFortmat), Constant.CSV_FILE_EXTENSION)
                CsvFileHelper.createFileCSV(settingInfo.Item(Constant.ErrorCsvFolder), csvErrorName, csvRecordErrorList)
            End If
        Next

    End Sub
End Class
