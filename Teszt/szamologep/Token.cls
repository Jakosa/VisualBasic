VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tkntype As String
Public number As Double
Public str As String

Private Sub Class_Initialize()
    tkntype = TKN_UNKNOWN
    number = 0#
    str = ""
End Sub

Public Sub Init(tt As String, d As Double, s As String)
    tkntype = tt
    number = d
    str = s
End Sub
