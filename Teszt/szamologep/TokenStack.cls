VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TokenStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim a() As New Token
Public inptr As Integer
Dim outptr As Integer

Private Sub Class_Initialize()
    ReDim a(0 To MAXTOKENS) As New Token
    inptr = 0
    outptr = 0
End Sub

Public Function Ures() As Boolean
    Ures = inptr <= 0
End Function

Public Function full() As Boolean
    full = inptr <= MAXTOKENS
End Function

Public Sub Lista() ' Put
    Dim i As Long
    Debug.Print "Print.Start"
    For i = LBound(a) To inptr
        Debug.Print a(i).str
    Next i
    Debug.Print "Print.End"
End Sub

Public Sub Tesz(t As Token) ' Put
    Set a(inptr) = t
    'Debug.Print "áélklksd: " & a(inptr).number
    inptr = inptr + 1
End Sub

Public Function Kap() As Token 'Get
    'Debug.Print "áélklksd: " & a(inptr).number
    'Set a(inptr) = Nothing
    'Set a(inptr) = New Token
    inptr = inptr - 1
    Set Kap = a(inptr)
End Function

Public Function Kap2(i As Integer) As Token 'Get
    Set Kap2 = a(i)
End Function

Public Function Keresgel() As Token ' Seek
    Set Keresgel = a(inptr - 1)
End Function
