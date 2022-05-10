Imports MyStudenApp.WebReference
Public Class ErrorDAO
    Public Shared Function CreateError(ByVal fileName As String, ByVal rowNumber As String, ByVal content As String) As Error__c
        Dim objError As New Error__c
        objError.Name = String.Format(Constant.ErrorMessage, fileName, rowNumber)
        objError.Content__c = content
        objError.Date__c = Date.Now
        objError.Date__cSpecified = True
        Return objError
    End Function

    Public Shared Function InsertError(ByVal posDataList As List(Of Error__c))
        If IsNothing(posDataList) OrElse Not posDataList.Any Then
            Return New List(Of UpsertResult)
        End If
        Return sfdcService.insert(posDataList.ToArray())
    End Function
End Class
