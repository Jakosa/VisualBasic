Attribute VB_Name = "Utils"
Option Explicit

Public Config As New Config

Public Function IsBoolean(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    IsBoolean = CBool(value)
ErrorHandler:
    ' Hamis lesz az érték
End Function

Public Function IsByte(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    IsByte = CByte(value)
ErrorHandler:
    ' Hamis lesz az érték
End Function
