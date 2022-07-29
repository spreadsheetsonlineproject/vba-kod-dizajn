Attribute VB_Name = "mdlSourceReport"
Option Explicit

Public Function getDataSheet(ByRef wb As Workbook) As Worksheet
    Set getDataSheet = wb.Worksheets(1)
End Function

Public Function getSourceData(ByRef source_sh As Worksheet) As Variant
    getSourceData = source_sh.Range("A1").CurrentRegion.Value
End Function

Public Function formatSource(ByRef data_sh As Worksheet)
    Dim target As Range
    
    Set target = data_sh.Range("A1").EntireColumn
    Call mdlFormat.setNumberFormat(target, "m/d/yyyy")
    
    Set target = data_sh.Range("B1").EntireColumn
    Call mdlFormat.setNumberFormat(target, "@")
    
    Set target = data_sh.Range("C1").EntireColumn
    Call mdlFormat.setNumberFormat(target, "_-* #,##0 [$Ft-hu-HU]_-;-* #,##0 [$Ft-hu-HU]_-;_-* ""-""?? [$Ft-hu-HU]_-;_-@_-")
End Function
