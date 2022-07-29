Attribute VB_Name = "mdlOriginal"
Option Explicit

Public Sub mainOriginal()


    ThisWorkbook.Sheets(1).Range("A1").EntireColumn.NumberFormat = "m/d/yyyy"
    
    ThisWorkbook.Sheets(1).Range("B1").EntireColumn.NumberFormat = "@"
    
    ThisWorkbook.Sheets(1).Range("C1").EntireColumn.NumberFormat = _
        "_-* #,##0 [$Ft-hu-HU]_-;-* #,##0 [$Ft-hu-HU]_-;_-* ""-""?? [$Ft-hu-HU]_-;_-@_-"
    

    Dim data_arr As Variant
    data_arr = ThisWorkbook.Sheets(1).Range("A1").CurrentRegion.Value
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lv As Long
    For lv = LBound(data_arr, 1) + 1 To UBound(data_arr, 1)
        dict(data_arr(lv, 2)) = dict(data_arr(lv, 2)) + data_arr(lv, 3)
    Next lv
    
    ThisWorkbook.Sheets(2).Cells(1, 1).Value = "Vásárló"
    ThisWorkbook.Sheets(2).Cells(1, 2).Value = "Összeg"
    
    Dim sh_lv As Long: sh_lv = 2
    Dim key As Variant
    For Each key In dict
        ThisWorkbook.Sheets(2).Cells(sh_lv, 1) = key
        ThisWorkbook.Sheets(2).Cells(sh_lv, 2) = dict(key)
        
        sh_lv = sh_lv + 1
    Next key
    
    ThisWorkbook.Sheets(2).Range("A1").EntireColumn.NumberFormat = "@"
    
    ThisWorkbook.Sheets(2).Range("B1").EntireColumn.NumberFormat = _
        "_-* #,##0 [$Ft-hu-HU]_-;-* #,##0 [$Ft-hu-HU]_-;_-* ""-""?? [$Ft-hu-HU]_-;_-@_-"

End Sub
