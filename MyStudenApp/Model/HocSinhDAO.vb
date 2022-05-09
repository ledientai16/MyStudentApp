Imports MyStudenApp.WebReference

Imports System.Reflection

Public Class HocSinhDAO

    Public Shared Function createHocSinh(ByVal csvRow As String()) As HocSinh__c
        Dim objHocSinh As New HocSinh__c
        objHocSinh.Name = CDbl(csvRow(ConstantCSV.MaHS))

        objHocSinh.Name = csvRow(ConstantCSV.Name)
        objHocSinh.HoHocSinh__c = csvRow(ConstantCSV.Ho)
        objHocSinh.GioiTinh__c = CBool(csvRow(ConstantCSV.Sex))
        objHocSinh.GioiTinh__cSpecified = True
        objHocSinh.NgaySinh__c = Date.Parse(csvRow(ConstantCSV.Birthday))
        objHocSinh.NgaySinh__cSpecified = True
        objHocSinh.Lop__c = csvRow(ConstantCSV.Class_ID)
        objHocSinh.diem1__c = CDbl(csvRow(ConstantCSV.Score1))
        objHocSinh.diem1__cSpecified = True
        objHocSinh.diem2__c = CDbl(csvRow(ConstantCSV.Score2))
        objHocSinh.diem2__cSpecified = True
        objHocSinh.diem3__c = CDbl(csvRow(ConstantCSV.Score3))
        objHocSinh.diem3__cSpecified = True
        Return objHocSinh
    End Function

    Public Shared Function upsertHocSinh(ByVal posDataList As List(Of HocSinh__c)) As List(Of UpsertResult)
        If IsNothing(posDataList) OrElse Not posDataList.Any Then
            Return New List(Of UpsertResult)
        End If
        Return sfdcService.upsert(posDataList.ToArray(), Constant.ExternalFieldUpsert)
    End Function
End Class
