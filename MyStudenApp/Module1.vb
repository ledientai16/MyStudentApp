Imports Ksvc.Utility

Module Module1
    Public sfdcService As SFDCHelper
    Public ReadOnly sysSetting As New SystemSettingHelper
    Public isSuccess As Boolean = False
    Public headerCSV As String
    Public csvFileError As String = ""
    Public errorArr As New List(Of String)
    Public loginInfo As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public settingInfo As Dictionary(Of String, String) = New Dictionary(Of String, String)



    Sub Main()
        MainController
    End Sub

End Module
