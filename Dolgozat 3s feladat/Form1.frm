VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim T(1 To 10) As Long
    Dim i As Long, hely As Long, ujelem As Long

    List1.Clear
    ujelem = 100

    For i = LBound(T) To UBound(T) - 1
        If Not (i = LBound(T)) Then
            T(i) = T(i - 1) + 10
        Else
            T(i) = 11
        End If

        List1.AddItem "Eredeti " & i & ": " & T(i)
    Next i

    hely = Linkeres(T, ujelem)
    Beilleszt T, hely, ujelem
    List1.AddItem ""

    For i = LBound(T) To UBound(T)
        List1.AddItem "Új " & i & ": " & T(i)
    Next i
End Sub

Private Function Linkeres(T() As Long, ujelem As Long) As Long
    Dim i As Long, index As Long
    index = -1

    For i = LBound(T) To UBound(T) - 1
        If i > LBound(T) Then
            If T(i - 1) <= ujelem And ujelem <= T(i) Then
                index = i
                Exit For
            Else
                If i = UBound(T) - 1 And index = -1 Then
                    index = UBound(T)
                End If
            End If
        Else
            If T(i) > ujelem Then
                index = LBound(T)
                Exit For
            End If
        End If
    Next i

    Linkeres = index
End Function

Private Sub Beilleszt(T() As Long, hely As Long, ujelem As Long)
    Dim i As Long, old As Long
    old = ujelem

    For i = hely To UBound(T)
        ujelem = T(i)
        T(i) = old
        old = ujelem
    Next i
End Sub
