Imports MyStudenApp.WebReference

''' <summary>
''' Sforce用処理クラス
''' </summary>
''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
Public Class SFDCHelper
#Region "メンバ"
    Public m_Sforce As SforceService
    Private m_User As String ' ユーザID
    Private m_Password As String  ' パスワード
    Private m_SecurityToken As String ' セキュリティートークン
    Private m_ProxyHost As String ' プロキシホスト
    Private m_ProxyPort As String ' プロキシポート
    Private m_ProxyUser As String ' プロキシユーザ
    Private m_ProxyPass As String ' プロキシパスワード
    Public m_Url As String ' 

    Private Const UPD_MAX_CNT As Integer = 200
    Private Const m_RetryCnt As Integer = 5 ' 例外時リトライ回数
    Private Const m_SleepTime As Integer = 1000 ' 待機時間（ミリ秒）
    Private Const m_SleepExTime As Integer = 2000 ' 例外時待機時間（ミリ秒）
    Private Const m_LoginTimeout As Integer = 10000000  ' ログインタイムアウト値（ミリ秒）
#End Region

#Region "コンストラクタ"

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="user">ユーザID</param>
    ''' <param name="password">パスワード</param> 
    ''' <param name="securityToken">セキュリティートークン</param>
    ''' <param name="proxyHost">プロキシホスト</param>
    ''' <param name="proxyPort">プロキシポート</param>
    ''' <param name="proxyUser">プロキシユーザ</param>
    ''' <param name="proxyPass">プロキシパスワード</param>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Sub New(ByVal user As String,
                   ByVal password As String,
                   ByVal securityToken As String,
                   ByVal proxyHost As String,
                   ByVal proxyPort As String,
                   ByVal proxyUser As String,
                   ByVal proxyPass As String)
        Me.m_User = user ' ユーザID
        Me.m_Password = password ' パスワード
        Me.m_SecurityToken = securityToken ' セキュリティートークン
        Me.m_ProxyHost = proxyHost ' プロキシホスト
        Me.m_ProxyPort = proxyPort ' プロキシポート
        Me.m_ProxyUser = proxyUser ' プロキシユーザ
        Me.m_ProxyPass = proxyPass ' プロキシパスワード
        Me.m_Sforce = New SforceService
        Me.m_Url = m_Sforce.Url
    End Sub

#End Region

#Region "ログイン"

    ''' <summary>
    ''' ログイン
    ''' </summary>
    ''' <returns>ログイン結果</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function login() As String
        Dim isLoginTimeout As Boolean = False
        Try
            Dim timeStamp As Date = m_Sforce.getServerTimestamp.timestamp
        Catch ex As Exception
            isLoginTimeout = True
        End Try

        If Not isLoginTimeout Then
            Return True
        End If

        Try
            'Timeout設定
            Me.m_Sforce.Timeout = m_LoginTimeout
            If Not String.IsNullOrEmpty(m_ProxyHost) And Not String.IsNullOrEmpty(m_ProxyPort) Then
                m_Sforce.Proxy = New Net.WebProxy(m_ProxyHost, CInt(m_ProxyPort))
            End If
            If Not String.IsNullOrEmpty(m_ProxyUser) And Not String.IsNullOrEmpty(m_ProxyPass) Then
                m_Sforce.Proxy.Credentials = New System.Net.NetworkCredential(m_ProxyUser, m_ProxyPass)
            End If
            Dim loginResult As LoginResult = Me.m_Sforce.login(Me.m_User, Me.m_Password + Me.m_SecurityToken)
            Me.m_Sforce.SessionHeaderValue = New SessionHeader
            Me.m_Sforce.SessionHeaderValue.sessionId = loginResult.sessionId
            Me.m_Sforce.Url = loginResult.serverUrl
            MainController.strExecutionUser = loginResult.userId
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#End Region

#Region "DMLのヘルパー"

    ''' <summary>
    ''' query
    ''' </summary>
    ''' <param name="strQueryString"></param>
    ''' <returns></returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function query(ByVal strQueryString As String) As QueryResult
        '再ログイン
        login()
        Return Me.m_Sforce.query(strQueryString)
    End Function
    ''' <summary>
    ''' queryAll
    ''' </summary>
    ''' <param name="strQueryString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function queryAll(ByVal strQueryString As String) As QueryResult
        '再ログイン
        login()
        Return Me.m_Sforce.queryAll(strQueryString)
    End Function

    ''' <summary>
    ''' queryMore
    ''' </summary>
    ''' <param name="strQueryLocator"></param>
    ''' <returns></returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function queryMore(ByVal strQueryLocator As String) As QueryResult
        Return Me.m_Sforce.queryMore(strQueryLocator)
    End Function

    ''' <summary>
    ''' 挿入する
    ''' </summary>
    ''' <param name="sobj">sobj</param>
    ''' <returns>ブール値</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function insert(ByVal sobj As sObject) As SaveResult
        Return insert(New sObject() {sobj})(0)
    End Function

    ''' <summary>
    ''' 挿入する
    ''' </summary>
    ''' <param name="sobjArray">sobjアレイ</param>
    ''' <returns>ブール値のリスト</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function insert(ByVal sobjArray() As sObject) As List(Of SaveResult)
        Dim iRetryCnt As Integer = 0
        Dim intCntSend As Integer = 0
        Dim Success As New List(Of Boolean)
        Dim listSaveResult As New List(Of SaveResult)
        'ログイン   
        login()
        Dim intCnt As Integer = 0
        Dim intLen As Integer = 0
        While intCnt < sobjArray.Length
            If sobjArray.Length - intCnt < Constant.NumRecordsCommit Then
                intLen = sobjArray.Length - intCnt
            Else
                intLen = Constant.NumRecordsCommit
            End If

            Dim sObj(intLen - 1) As sObject
            Array.Copy(sobjArray, intCnt, sObj, 0, intLen)

            '登録
            Dim saveResult() As SaveResult = Me.m_Sforce.create(sObj)
            'Dim saveResult() As SaveResult = Me.m_Sforce.upsert(sObj)

            listSaveResult.AddRange(saveResult)
            For i As Integer = 0 To saveResult.Length - 1
                If saveResult(i).success Then
                    ' 更新リストに追加
                    sObj(i).Id = saveResult(i).id
                    errorArr.Add("")
                Else
                    errorArr.Add(saveResult(i).errors(0).message.ToString)
                End If
                Success.Add(saveResult(i).success)
            Next
            intCnt += intLen

            '*** 1000件更新ごとにSleep実施 ***
            intCntSend += Constant.NumRecordsCommit
            If intCntSend >= m_SleepTime Then
                System.Threading.Thread.Sleep(m_SleepTime)
                intCntSend = 0
            End If

            '*** Retry回数初期化 ***
            iRetryCnt = 0
        End While
        Return listSaveResult
    End Function

    ''' <summary>
    ''' 削除する
    ''' </summary>
    ''' <param name="sobj">sobj</param>
    ''' <returns>ブール値</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function delete(ByVal sobj As sObject) As DeleteResult
        Return delete(New sObject() {sobj})(0)
    End Function

    ''' <summary>
    ''' 削除する
    ''' </summary>
    ''' <param name="sobjArray">sobjアレイ</param>
    ''' <returns>ブール値のリスト</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function delete(ByVal sobjArray() As sObject) As List(Of DeleteResult)
        Dim iRetryCnt As Integer = 0
        Dim intCntSend As Integer = 0
        Dim Success As New List(Of Boolean)
        Dim listDeleteResult As New List(Of DeleteResult)
        'ログイン   
        login()

        Dim intCnt As Integer = 0
        Dim intLen As Integer = 0
        While intCnt < sobjArray.Length
            If sobjArray.Length - intCnt < Constant.NumRecordsCommit Then
                intLen = sobjArray.Length - intCnt
            Else
                intLen = Constant.NumRecordsCommit
            End If

            Dim sObj(intLen - 1) As sObject
            Dim delIdList(intLen - 1) As String
            Array.Copy(sobjArray, intCnt, sObj, 0, intLen)
            For i As Long = 0 To sObj.Length - 1
                delIdList(i) = sObj(i).Id
            Next

            '登録
            Dim deleteResult() As DeleteResult = Me.m_Sforce.delete(delIdList)
            listDeleteResult.AddRange(deleteResult)
            For i As Integer = 0 To deleteResult.Length - 1
                If deleteResult(i).success Then
                    ' 更新リストに追加
                    errorArr.Add("")
                Else
                    errorArr.Add(deleteResult(i).errors(0).message.ToString)
                End If
                Success.Add(deleteResult(i).success)
            Next
            intCnt += intLen

            '*** 1000件更新ごとにSleep実施 ***
            intCntSend += Constant.NumRecordsCommit
            If intCntSend >= m_SleepTime Then
                System.Threading.Thread.Sleep(m_SleepTime)
                intCntSend = 0
            End If

            '*** Retry回数初期化 ***
            iRetryCnt = 0
        End While
        Return listDeleteResult
    End Function


    Public Function emptyRecycleBin(ByVal sobjArray() As sObject) As List(Of EmptyRecycleBinResult)
        Dim iRetryCnt As Integer = 0
        Dim intCntSend As Integer = 0
        Dim Success As New List(Of Boolean)
        Dim listDeleteResult As New List(Of EmptyRecycleBinResult)
        'ログイン   
        login()

        Dim intCnt As Integer = 0
        Dim intLen As Integer = 0
        While intCnt < sobjArray.Length
            If sobjArray.Length - intCnt < Constant.NumRecordsCommit Then
                intLen = sobjArray.Length - intCnt
            Else
                intLen = Constant.NumRecordsCommit
            End If

            Dim sObj(intLen - 1) As sObject
            Dim delIdList(intLen - 1) As String
            Array.Copy(sobjArray, intCnt, sObj, 0, intLen)
            For i As Long = 0 To sObj.Length - 1
                delIdList(i) = sObj(i).Id
            Next

            '登録
            Dim deleteResult() As EmptyRecycleBinResult = Me.m_Sforce.emptyRecycleBin(delIdList)
            listDeleteResult.AddRange(deleteResult)
            For i As Integer = 0 To deleteResult.Length - 1
                If deleteResult(i).success Then
                    ' 更新リストに追加
                    errorArr.Add("")
                Else
                    errorArr.Add(deleteResult(i).errors(0).message.ToString)
                End If
                Success.Add(deleteResult(i).success)
            Next
            intCnt += intLen

            '*** 1000件更新ごとにSleep実施 ***
            intCntSend += Constant.NumRecordsCommit
            If intCntSend >= m_SleepTime Then
                System.Threading.Thread.Sleep(m_SleepTime)
                intCntSend = 0
            End If

            '*** Retry回数初期化 ***
            iRetryCnt = 0
        End While
        Return listDeleteResult
    End Function

    ''' <summary>
    ''' 更新する
    ''' </summary>
    ''' <param name="sobj">sobj</param>
    ''' <returns>ブール値</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function update(ByVal sobj As sObject) As Boolean
        Return update(New sObject() {sobj})(0)
    End Function


    ''' <summary>
    ''' 更新する
    ''' </summary>
    ''' <param name="sobjArray">sobjアレイ</param>
    ''' <returns>ブール値のリスト</returns>
    ''' <remarks>Create: 2014/02/18 kvc thuy</remarks>
    Public Function update(ByVal sobjArray() As sObject) As List(Of Boolean)
        Dim iRetryCnt As Integer = 0
        Dim intCntSend As Integer = 0
        Dim Success As New List(Of Boolean)

        'ログイン   
        login()

        Dim intCnt As Integer = 0
        Dim intLen As Integer = 0
        While intCnt < sobjArray.Length
            If sobjArray.Length - intCnt < Constant.NumRecordsCommit Then
                intLen = sobjArray.Length - intCnt
            Else
                intLen = UPD_MAX_CNT
            End If

            Dim sObj(intLen - 1) As sObject
            Array.Copy(sobjArray, intCnt, sObj, 0, intLen)

            '登録
            Dim saveResult As SaveResult() = Me.m_Sforce.update(sObj)

            For i As Integer = 0 To saveResult.Length - 1
                If saveResult(i).success Then
                    ' 更新リストに追加
                    sObj(i).Id = saveResult(i).id
                Else
                    errorArr.Add((saveResult(i).errors(0).message.ToString))
                End If
                Success.Add(saveResult(i).success)
            Next
            intCnt += intLen

            '*** 1000件更新ごとにSleep実施 ***
            intCntSend += Constant.NumRecordsCommit
            If intCntSend >= m_SleepTime Then
                System.Threading.Thread.Sleep(m_SleepTime)
                intCntSend = 0
            End If

            '*** Retry回数初期化 ***
            iRetryCnt = 0
        End While
        Return Success
    End Function
    ''' <summary>
    ''' サートする
    ''' </summary>
    ''' <param name="sobjArray">sobjアレイ</param>
    ''' <returns>UpsertResultのリスト</returns>
    ''' <remarks>Create: 2014/05/30 ksvc vinh</remarks>
    Public Function upsertReturnUpsertResult(ByVal sobjArray() As sObject, ByVal externalField As String) As List(Of UpsertResult)
        Dim upsertResultList As New List(Of UpsertResult)
        'ログイン   
        login()
        upsertResultList.AddRange(m_Sforce.upsert(externalField, sobjArray.ToArray()))
        Return upsertResultList
    End Function
    ''' <summary>
    ''' サートする
    ''' </summary>
    ''' <param name="sobjArray">sobjアレイ</param>
    ''' <returns>ブール値のリスト</returns>
    ''' <remarks>Create: 2014/06/02 Ksvc Vinh Hua Quoc</remarks>
    Public Function upsert(ByVal sobjArray() As sObject, ByVal externalField As String) As List(Of UpsertResult)
        Dim iRetryCnt As Integer = 0
        Dim intCntSend As Integer = 0
        Dim Success As New List(Of Boolean)
        Dim upsertResultList As New List(Of UpsertResult)
        'ログイン
        login()
        Dim intCnt As Integer = 0
        Dim intLen As Integer = 0
        While intCnt < sobjArray.Length
            If sobjArray.Length - intCnt < Constant.NumRecordsCommit Then
                intLen = sobjArray.Length - intCnt
            Else
                intLen = Constant.NumRecordsCommit
            End If

            Dim sObj(intLen - 1) As sObject
            Array.Copy(sobjArray, intCnt, sObj, 0, intLen)

            '登録
            Dim upsertResult() As UpsertResult = Me.m_Sforce.upsert(externalField, sObj)
            upsertResultList.AddRange(upsertResult)

            For i As Integer = 0 To upsertResult.Length - 1
                If upsertResult(i).success Then
                    ' 更新リストに追加
                    sObj(i).Id = upsertResult(i).id
                Else
                    errorArr.Add((upsertResult(i).errors(0).message.ToString))
                End If
                Success.Add(upsertResult(i).success)
            Next
            intCnt += intLen

            '*** 1000件更新ごとにSleep実施 ***
            intCntSend += Constant.NumRecordsCommit
            If intCntSend >= m_SleepTime Then
                System.Threading.Thread.Sleep(m_SleepTime)
                intCntSend = 0
            End If

            '*** Retry回数初期化 ***
            iRetryCnt = 0
        End While
        Return upsertResultList
    End Function
#End Region
End Class
