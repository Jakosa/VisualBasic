VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ellenõrzés!"
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Megadandó szám:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Prim szám-e?"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PeldaSzoveg = "Írjon be egy számot."

Private Sub Form_Load()
    Text1.Text = PeldaSzoveg
End Sub

Private Sub Command1_Click()
    Dim szam As Long

    If Trim(Text1.Text) = "" Then
        MsgBox "Nincs megadva szám (Üres a cella vagy szóközzel, tabulátorral stb van tele.)!", vbCritical, "Hiba!"
        Exit Sub
    End If

    If IsLong(Text1.Text) Then
        szam = Text1.Text
    Else
        MsgBox "Nem long típusú szám lett megadva!", vbCritical, "Hiba!"
        Exit Sub
    End If

    If PrimE(szam) Then
        Text2.Text = "Prim szám!"
    Else
        Text2.Text = "Nem prim szám!"
    End If

    AlattLevoPrimek szam
End Sub

Private Sub Text1_GotFocus()
    If Text1.Text = PeldaSzoveg Then
        Text1.Text = ""
    End If
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = "" Then
        Text1.Text = PeldaSzoveg
    End If
End Sub

Private Function PrimE(x As Long) As Boolean
    If x = 1 Or x = 0 Then
        Exit Function
    End If

    If x = 2 Then
        PrimE = True
        Exit Function
    End If

    If x Mod 2 = 0 Then
        Exit Function
    End If

    Dim p As Boolean, i As Long
    p = True

    For i = 3 To CLng(Sqr(x)) Step 2
        If x Mod i = 0 Then
            p = False
            Exit For
        End If
    Next i

    PrimE = p
End Function

Private Sub AlattLevoPrimek(szam As Long)
    Dim i As Long
    List1.Clear

    If szam > 2 Then
        For i = 2 To szam - 1
            If PrimE(i) Then
                List1.AddItem i
            End If
        Next i
    ElseIf szam = 2 Then
        List1.AddItem 2
    Else
        List1.AddItem "Kisebb mint 2 így bíztos nincsenek prím számok amiket ki lehetne írni."
    End If
End Sub

Private Function IsLong(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    Dim i As Boolean
    i = CLng(value)
    IsLong = True
ErrorHandler:
    ' Hamis lesz az érték
End Function


