VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TokenQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim a() As New Token
Dim inptr As Integer
Dim outptr As Integer

Private Sub Class_Initialize()
    ReDim a(0 To MAXTOKENS) As New Token
    inptr = 0
    outptr = 0
End Sub

Public Function Ures() As Boolean
    Ures = inptr <= outptr
End Function

Public Function full() As Boolean
    full = inptr <= MAXTOKENS
End Function

Public Sub Lista() ' Put
    Dim i As Long
    Debug.Print "Print0.Start"
    For i = LBound(a) To inptr
        Debug.Print a(i).number
    Next i
    Debug.Print "Print0.End"
End Sub

Public Sub Tesz(t As Token)
    Set a(inptr) = t
    inptr = inptr + 1
End Sub

Public Function Kap() As Token
    Set Kap = a(outptr)
    outptr = outptr + 1
End Function

Public Sub Takaritas()
    Dim i As Long

    For i = LBound(a) To inptr
        If a(i).str = "" And a(i).tkntype = TKN_UNKNOWN Then
            Debug.Print "baj van " & i
        End If
    Next i
End Sub
