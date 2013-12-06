Attribute VB_Name = "EnumSetWindowPos"
' Fejl�c
' K�sz�tette: Jakosa Csaba �rp�d
' Fejl�c v�ge

Option Explicit

Public Enum EnSetWindowPos
    SWP_FRAMECHANGED = &H20
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum
