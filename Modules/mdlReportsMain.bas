Attribute VB_Name = "mdlReportsMain"
Option Explicit

Public Sub sumReportMain()

    Rem get source sheet
    Dim data_sh As Worksheet
    Set data_sh = mdlSourceReport.getDataSheet(ThisWorkbook)
    
    Rem format surce sheet
    Call mdlSourceReport.formatSource(data_sh)
    
    Rem get source data
    Dim data_arr As Variant
    data_arr = mdlSourceReport.getSourceData(data_sh)
    
    Rem create sum report
    Dim sum_report_sh As Worksheet
    Set sum_report_sh = mdlSumReport.getSumReportSheet(ThisWorkbook)
    Call mdlSumReport.createSumReport(sum_report_sh, data_arr)

End Sub
