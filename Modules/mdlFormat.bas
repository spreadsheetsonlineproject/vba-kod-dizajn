Attribute VB_Name = "mdlFormat"
Option Explicit

Public Function setNumberFormat(ByRef target As Range, ByVal format As String)
    target.NumberFormat = format
End Function
