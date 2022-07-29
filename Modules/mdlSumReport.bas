Attribute VB_Name = "mdlSumReport"
Option Explicit

Private Const sum_report_col_number As Integer = 2

Public Function getSumReportSheet(ByRef wb As Workbook) As Worksheet
    Set getSumReportSheet = wb.Worksheets(2)
End Function

Public Sub createSumReport(ByRef report_sh As Worksheet, ByRef data_arr As Variant)
    Dim sum_dict As Object
    Set sum_dict = sumReportData(data_arr)
    
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
    Dim tmp_arr(1 To sum_report_col_number) As String
    tmp_arr(1) = "Vásárló"
    tmp_arr(2) = "Összeg"
    sumReportHeader = tmp_arr
End Function

Private Function addSumReportData(ByRef report_sh As Worksheet, ByRef dict As Object)
    Dim sh_lv As Long: sh_lv = 2
    Dim key As Variant
    Dim tmp_row(1 To sum_report_col_number) As Variant
    For Each key In dict
        tmp_row(1) = key
        tmp_row(2) = dict(key)
        Call mdlWriteReports.addReportRow(report_sh, tmp_row, sh_lv)
        
        sh_lv = sh_lv + 1
    Next key
End Function
