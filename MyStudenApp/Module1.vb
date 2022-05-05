Imports Ksvc.Utility

Module Module1
    Public ReadOnly sysSetting As New SystemSettingHelper
    Public isSuccess As Boolean = False
    Public headerCSV As String
    Public csvFileError As String = ""
    Public errorArr As New List(Of String)
    Sub Main()
        MainController
    End Sub

End Module
