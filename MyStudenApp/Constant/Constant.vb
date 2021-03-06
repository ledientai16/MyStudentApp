Imports Ksvc.Utility

Public Class Constant
#Region "dd"
    Public Const SystemSettingFile As String = "systemsetting.xml"
    Public Const LoginInfo As String = "logininfor.xml"
    Public Shared ReadOnly RootDirectory As String = Common.getAbsPath()
    Public Const TEXT_ENCODING As String = "Shift-JIS"

    Public Const LogFolderName = "LogFolder"
    Public Const LogErrorFileName = "Error_{0:yyyyMMddHHmmss}.log"
    Public Const LogExecuteFileName = "Execute_{0:yyyyMMddHHmmss}.log"


    Public Const DateTimeFormat As String = "yyyyMMddHHmm"
    Public Const CultureInfo As String = "en_Us-EN"

    Public Const CSV_FILE_EXTENSION As String = ".csv"
    Public Const CSV_Delimeter As String = ","
    Public Const LoginInfoFile As String = "LoginInfo.xml"

    Public Const ErrorMessage As String = "File {0} row {1} error"

    Public Const ExternalFieldUpsert As String = "MaHS__c"

    Public Const CSVFileNameMsg As String = "CSV file name :{0}"

    Public Const CSVSuccessFileName As String = "List_Success_{0}{1}"
    Public Const CSVErrorFileName As String = "List_Error_{0}{1}"

    Public Const DateTimeFortmat As String = "yyyyMMddHHmm"

    Public Const ExitCodeSuccess As Integer = 0
    Public Const ExitCodeError As Integer = 1


#End Region

#Region "ss"
    Public Shared NumRecordsCommit As Integer = 200
    Public Const LengthIdWhereIn As Integer = 19500
#End Region

#Region "Login info tag"
    Public Const UserNameTag As String = "username"
    Public Const PassWordTag As String = "password"
    Public Const ProxyHostTag As String = "proxy_host"
    Public Const ProxyPortTag As String = "proxy_port"
    Public Const SecurityTokenTag As String = "securitytoken"
    Public Const ProxyUserTag As String = "proxy_username"
    Public Const ProxyPassTag As String = "proxy_password"
#End Region
#Region "Setting info tag"
    Public Const BackupCsvFoler As String = "BackupDataFolder"
    Public Const ErrorCsvFolder As String = "ErrorDataFolder"
    Public Const InputFolder As String = "InputFolder"
#End Region


End Class
