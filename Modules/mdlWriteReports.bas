Attribute VB_Name = "mdlWriteReports"
Option Explicit

Public Function addReportHeader(ByRef report_sh As Worksheet, ByRef header_arr As Variant)
    report_sh.Range("A1").Resize(1, UBound(header_arr)).Value = header_arr
End Function

Public Function addReportRow(ByRef report_sh As Worksheet, ByRef row_data As Variant, ByVal target_row As Long)
    report_sh.Range(report_sh.Cells(target_row, 1), report_sh.Cells(target_row, 1)).Resize(1, UBound(row_data)).Value = row_data
End Function

