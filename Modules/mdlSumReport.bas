Attribute VB_Name = "mdlSumReport"
Option Explicit

Private Const SUM_REPORT_COL_NUMBER As Integer = 2
Private Const SUM_REPORT_SHEET_NAME As String = "ertek riport"

Public Function getSumReportSheet(ByRef wb As Workbook) As Worksheet

    Dim c_sh As Object
    For Each c_sh In wb.Worksheets
        If c_sh.Name = SUM_REPORT_SHEET_NAME Then
            Set getSumReportSheet = c_sh
            Exit Function
        End If
    Next c_sh

    Set getSumReportSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    getSumReportSheet.Name = SUM_REPORT_SHEET_NAME
    
End Function

Public Sub createSumReport(ByRef report_sh As Worksheet, ByRef data_arr As Variant)
    Dim sum_dict As Object
    Set sum_dict = sumReportData(data_arr)
    
    Set sum_dict = sortData(sum_dict)
    
    Dim sum_header As Variant
    sum_header = sumReportHeader
    
    Call mdlWriteReports.addReportHeader(report_sh, sum_header)
    
    Call addSumReportData(report_sh, sum_dict)
    
    Dim target As Range
    Set target = report_sh.Range("A1").EntireColumn
    Call mdlFormat.setNumberFormat(target, "@")
    
    Set target = report_sh.Range("B1").EntireColumn
    Call mdlFormat.setNumberFormat(target, "_-* #,##0 [$Ft-hu-HU]_-;-* #,##0 [$Ft-hu-HU]_-;_-* ""-""?? [$Ft-hu-HU]_-;_-@_-")
End Sub

Private Function sortData(ByRef sum_dict As Object) As Object
    Dim sorted_dict As Object
    
    
    
    Set sortData = sorted_dict
End Function

Private Function sumReportData(ByRef data_arr As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lv As Long
    For lv = LBound(data_arr, 1) + 1 To UBound(data_arr, 1)
        dict(data_arr(lv, 2)) = dict(data_arr(lv, 2)) + data_arr(lv, 3)
    Next lv
    
    Set sumReportData = dict
End Function

Private Function sumReportHeader() As Variant
    Dim tmp_arr(1 To SUM_REPORT_COL_NUMBER) As String
    tmp_arr(1) = "V�s�rl�"
    tmp_arr(2) = "�sszeg"
    sumReportHeader = tmp_arr
End Function

Private Function addSumReportData(ByRef report_sh As Worksheet, ByRef dict As Object)
    Dim sh_lv As Long: sh_lv = 2
    Dim key As Variant
    Dim tmp_row(1 To SUM_REPORT_COL_NUMBER) As Variant
    For Each key In dict
        tmp_row(1) = key
        tmp_row(2) = dict(key)
        Call mdlWriteReports.addReportRow(report_sh, tmp_row, sh_lv)
        
        sh_lv = sh_lv + 1
    Next key
End Function
