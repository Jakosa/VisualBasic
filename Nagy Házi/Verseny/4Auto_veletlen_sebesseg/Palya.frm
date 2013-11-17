VERSION 5.00
Begin VB.Form Palya 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   Caption         =   "Palya"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_AutoLista 
      Interval        =   500
      Left            =   1320
      Top             =   9360
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Aut�k"
      Height          =   3495
      Left            =   9600
      TabIndex        =   7
      Top             =   2400
      Width           =   4815
      Begin VB.TextBox AutoListaText 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox AutoLista 
         Height          =   315
         ItemData        =   "Palya.frx":0000
         Left            =   240
         List            =   "Palya.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Versenyadatok"
      Height          =   3615
      Left            =   9600
      TabIndex        =   5
      Top             =   6120
      Width           =   4815
      Begin VB.TextBox VersenyAdatokText 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox VersenyAdatok 
         Height          =   315
         ItemData        =   "Palya.frx":0034
         Left            =   240
         List            =   "Palya.frx":004A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Timer Timer_VersenyAdatok 
      Interval        =   500
      Left            =   720
      Top             =   9360
   End
   Begin VB.Timer Timer_Korok 
      Interval        =   50
      Left            =   120
      Top             =   9360
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Szektor 2"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Szektor 1"
      Height          =   195
      Left            =   4680
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
   Begin VB.Line SzektorVonal 
      Index           =   2
      X1              =   1920
      X2              =   600
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line SzektorVonal 
      Index           =   1
      X1              =   6480
      X2              =   5040
      Y1              =   5880
      Y2              =   4920
   End
   Begin VB.Line SzektorVonal 
      Index           =   0
      X1              =   4560
      X2              =   3600
      Y1              =   360
      Y2              =   2280
   End
   Begin VB.Label KorKiiras 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "K�r�k sz�ma: 0/0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "C�l"
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   4560
      UseMnemonic     =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Start"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      UseMnemonic     =   0   'False
      Width           =   330
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   43
      X1              =   5040
      X2              =   4200
      Y1              =   5640
      Y2              =   6240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   42
      X1              =   4200
      X2              =   3360
      Y1              =   6240
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   41
      X1              =   2880
      X2              =   2040
      Y1              =   7080
      Y2              =   6720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   40
      X1              =   2040
      X2              =   720
      Y1              =   6720
      Y2              =   5520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   39
      X1              =   5400
      X2              =   4680
      Y1              =   6720
      Y2              =   6960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   38
      X1              =   5760
      X2              =   5400
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   37
      X1              =   4680
      X2              =   3720
      Y1              =   6960
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   36
      X1              =   2280
      X2              =   3360
      Y1              =   6000
      Y2              =   6480
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   35
      X1              =   1560
      X2              =   2280
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   34
      X1              =   3720
      X2              =   2880
      Y1              =   7200
      Y2              =   7080
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   33
      X1              =   5760
      X2              =   5400
      Y1              =   4200
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   32
      X1              =   5400
      X2              =   5040
      Y1              =   4920
      Y2              =   5640
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   31
      X1              =   6960
      X2              =   5760
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   30
      X1              =   7200
      X2              =   6120
      Y1              =   4920
      Y2              =   5280
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   29
      X1              =   6120
      X2              =   5760
      Y1              =   5280
      Y2              =   6000
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   2
      X1              =   7680
      X2              =   6960
      Y1              =   3720
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   28
      X1              =   8040
      X2              =   7200
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   27
      X1              =   8400
      X2              =   8040
      Y1              =   3840
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   26
      X1              =   8400
      X2              =   8400
      Y1              =   3240
      Y2              =   3840
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   25
      X1              =   7680
      X2              =   7680
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   0
      X1              =   7680
      X2              =   4560
      Y1              =   3120
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   24
      X1              =   8400
      X2              =   4200
      Y1              =   2640
      Y2              =   840
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   23
      X1              =   4560
      X2              =   3720
      Y1              =   1920
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   22
      X1              =   4200
      X2              =   3360
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   21
      X1              =   3720
      X2              =   3360
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   20
      X1              =   3360
      X2              =   3240
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   19
      X1              =   3240
      X2              =   3000
      Y1              =   1920
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   18
      X1              =   2400
      X2              =   1800
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   17
      X1              =   1800
      X2              =   1560
      Y1              =   3240
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   16
      X1              =   1560
      X2              =   1560
      Y1              =   4800
      Y2              =   5400
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   15
      X1              =   1560
      X2              =   1560
      Y1              =   3600
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   14
      X1              =   3000
      X2              =   2400
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   13
      X1              =   960
      X2              =   720
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   12
      X1              =   2040
      X2              =   1440
      Y1              =   2520
      Y2              =   2640
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   11
      X1              =   2640
      X2              =   2400
      Y1              =   1680
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   10
      X1              =   2400
      X2              =   2040
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   9
      X1              =   3360
      X2              =   2880
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   8
      X1              =   8400
      X2              =   8400
      Y1              =   2640
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   7
      X1              =   1440
      X2              =   960
      Y1              =   2640
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   6
      X1              =   1560
      X2              =   1560
      Y1              =   4200
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   5
      X1              =   2880
      X2              =   2640
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   4
      X1              =   720
      X2              =   720
      Y1              =   4320
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   3
      X1              =   720
      X2              =   720
      Y1              =   3720
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Index           =   1
      X1              =   720
      X2              =   720
      Y1              =   4920
      Y2              =   5520
   End
   Begin VB.Menu game 
      Caption         =   "J�t�k"
      Begin VB.Menu Start 
         Caption         =   "Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu Stop 
         Caption         =   "Stop"
         Shortcut        =   ^C
      End
      Begin VB.Menu gamebar1 
         Caption         =   "-"
      End
      Begin VB.Menu NewGame 
         Caption         =   "�j j�t�k"
         Shortcut        =   ^G
      End
      Begin VB.Menu SaveResult 
         Caption         =   "V�geredm�ny ment�se"
      End
      Begin VB.Menu gamebar2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Kilp�s"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "Be�ll�t�sok"
      Begin VB.Menu Nyomvonal 
         Caption         =   "Nyomvonal"
         Checked         =   -1  'True
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu Help 
      Caption         =   "S�g�"
      Begin VB.Menu About 
         Caption         =   "N�vjegy"
      End
   End
End
Attribute VB_Name = "Palya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Autok(1 To 4) As New Auto   ' Aut�k be�ll�t�s�t t�rol� t�mb.
Dim AutokSzama As Byte          ' Aut�k sz�ma
Dim Korok As Byte               ' �ppen h�nyadik k�rn�l tartunk.
Dim MKorokSzama As Byte         ' Maximum k�r�k sz�ma
Dim Started As Boolean          ' Jelzi hogy elindult-e m�r a j�t�k vagy sem.
Const KezdokorErteke = 1        ' T�rolja hogy mennyit�l induljon az els� k�r.
Const BorderWidth = 2           ' Aut�k vonal�nak sz�less�ge.
Const ex = 0.6
Const ey = -1

Public Property Get GetKezdokorErteke() As Byte
    GetKezdokorErteke = KezdokorErteke
End Property

Public Property Get GetMKorokSzama() As Byte
    GetMKorokSzama = MKorokSzama
End Property

Private Sub Form_Load()
    Started = False
    AutokSzama = 0
    Korok = KezdokorErteke
    MKorokSzama = 5
    SetKorokSzama Korok

    VersenyAdatok.ListIndex = 0
    AutoLista.ListIndex = 0
End Sub

Private Sub Form_Terminate()
    'Dispose_Game
End Sub

Private Sub New_Game(ASzama As Byte)
    If Started Then
        Exit Sub
    End If

    Dim i As Byte
    i = 1

    Dim T(1 To 4) As String
    T(1) = "piros"
    T(2) = "k�k"
    T(3) = "s�rga"
    T(4) = "z�ld"

    For i = 1 To ASzama
        Autok(i).Load i ' Bet�ltj�k �jk�nt a vonalat
        Autok(i).SetEX ex
        Autok(i).SetEY ey
        Autok(i).SetX0 1100 - i * 20
        Autok(i).SetY0 4000 - i * 100
        Autok(i).SetColor T(i) ' Ha kell sz�nez�s csak akkor.
        Autok(i).SetBorderWidth BorderWidth
        Autok(i).Show
    Next i

    AutokSzama = i - 1
End Sub

Private Sub Dispose_Game()
    If Started Then
        Exit Sub
    End If

    Dim i As Byte
    i = 1

    For i = 1 To AutokSzama
        Autok(i).Dispose
        Set Autok(i) = Nothing
    Next i

    AutokSzama = 0
End Sub

Private Sub NewGame_Click()
    Started = False
    Dispose_Game
    AutokSzama = 0
    Timer_AutoLista.Enabled = True
    AutoLista.Enabled = True
    Korok = KezdokorErteke
    MKorokSzama = 5
    SetKorokSzama Korok

    VersenyAdatok.ListIndex = 0
    AutoLista.ListIndex = 0
End Sub

Private Sub Nyomvonal_Click()
    If Nyomvonal.Checked Then
        Nyomvonal.Checked = False
    Else
        Nyomvonal.Checked = True
    End If

    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        Autok(i).SetNyomvonal Nyomvonal.Checked
    Next i
End Sub

Private Sub Start_Click()
    If AutokSzama = 0 Then
        MsgBox "M�g nincsenek kiv�lasztva aut�k!"
        Exit Sub
    End If
    If Not Started Then
        Started = True
    End If

    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        Autok(i).Start
    Next i
End Sub

Private Sub Stop_Click()
    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        Autok(i).Stop_Kocsi
    Next i
End Sub

Private Sub About_Click()
    AboutForm.Show
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub SetKorokSzama(KorSz As Byte)
    KorKiiras.Caption = "K�r�k sz�ma: " & KorSz & "/" & MKorokSzama
End Sub

Private Sub Timer_Korok_Timer()
    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        If Korok < Autok(i).GetKorokSzama Then
            Korok = Autok(i).GetKorokSzama

            If Korok > MKorokSzama Then
                Stop_Click
                MsgBox "V�ge a j�t�knak! Nyertes aut� sz�ma: " & i
                Exit Sub
            End If

            SetKorokSzama Korok
        End If
    Next i
End Sub

Private Sub Timer_VersenyAdatok_Timer()
    Dim i As Byte, szam As Single, szin As String

    Select Case VersenyAdatok.List(VersenyAdatok.ListIndex)
        Case "Aut�k sorrendje"
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If
        Case "Legjobb 1. szektor"
            SzektoridoKiiras 1
        Case "Legjobb 2. szektor"
            SzektoridoKiiras 2
        Case "Legjobb 3. szektor"
            SzektoridoKiiras 3
        Case "Legjobb k�rid�"
            Dim lkor As Byte, aszam As Byte
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            If Korok = KezdokorErteke Then
                AddVAText "Nincs m�g m�rt k�rid�!"
                Exit Sub
            End If

            For i = LBound(Autok) To AutokSzama
                If szam < Autok(i).GetLegjobbKorido Then
                    aszam = i
                    szin = Autok(i).GetColor
                    szam = Autok(i).GetLegjobbKorido
                    lkor = Autok(i).GetLegjobbKoridoSzama
                End If
            Next i

            AddVAText "Legjobb k�r ideje: " & szam & " m�sodperc"
            AddVAText "Els� szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(1) & " m�sodperc"
            AddVAText "M�sodik szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(2) & " m�sodperc"
            AddVAText "Harmadik szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(3) & " m�sodperc"
            AddVAText ""
            AddVAText "Az id� a(z) " & lkor & ". k�rben ker�lt be�ll�t�sra."
            AddVAText "A(z) id�t be�ll�totta a " & szin & " szin� aut�."
        Case "Elm�leti legjobb k�rid�"
            Dim T(1 To 3) As Single, TSzin(1 To 3) As String, ljekorido As Single
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            If Korok = KezdokorErteke Then
                AddVAText "Nincs m�g m�rt k�rid�!"
                Exit Sub
            End If

            For i = 1 To 3
                T(i) = LegjobbSzektorido(i, TSzin(i))
                ljekorido = ljekorido + T(i)
            Next i

            AddVAText "Legjobb elm�leti k�ri�: " & ljekorido & " m�sodperc"
            AddVAText "Els� szektor ideje: " & T(1) & " m�sodperc"
            AddVAText "Aut� sz�ne: " & TSzin(1)
            AddVAText ""
            AddVAText "M�sodik szektor ideje: " & T(2) & " m�sodperc"
            AddVAText "Aut� sz�ne: " & TSzin(2)
            AddVAText ""
            AddVAText "Harmadik szektor ideje: " & T(3) & " m�sodperc"
            AddVAText "Aut� sz�ne: " & TSzin(3)
        Case Else
            CleanVAText
            AddVAText "Hiba!"
    End Select
End Sub

Private Function LegjobbSzektorido(ByVal a As Integer, szin As String) As Single
    Dim i As Byte, szam As Single
    CleanVAText

    For i = LBound(Autok) To AutokSzama
        If szam < Autok(i).GetLegjobbSzektoridok(a) Then
            szin = Autok(i).GetColor
            szam = Autok(i).GetLegjobbSzektoridok(a)
        End If
    Next i

    LegjobbSzektorido = szam
End Function

Private Sub SzektoridoKiiras(a As Integer)
    Dim i As Byte, szam As Single, szin As String
    CleanVAText

    If Not Started Then
        NoStartedGameVAText
        Exit Sub
    End If

    szam = LegjobbSzektorido(a, szin)

    If szam = 10000000 Then ' Kezd��rt�k
        AddVAText "Nincs m�g m�rt szektorid�!"
        Exit Sub
    End If

    AddVAText "Legjobb szektorid�: " & szam & " m�sodperc"
    AddVAText "Az id�t be�ll�totta a " & szin & " szin� aut�."
End Sub

Private Sub NoStartedGameVAText()
    VersenyAdatokText.Text = "M�g nem indult el a j�t�k!"
End Sub

Private Sub CleanVAText()
    VersenyAdatokText.Text = ""
End Sub

Private Sub AddVAText(Szoveg As String)
    VersenyAdatokText.Text = VersenyAdatokText.Text & Szoveg & vbCrLf
End Sub

Private Sub Timer_AutoLista_Timer()
    If Started Then
        AutoLista.Enabled = False
        Timer_AutoLista.Enabled = False
    End If

    Select Case AutoLista.List(AutoLista.ListIndex)
        Case "Kett� aut�"
            Dispose_Game
            New_Game 2
            AutokKiirasa
        Case "H�rom aut�"
            Dispose_Game
            New_Game 3
            AutokKiirasa
        Case "N�gy aut�"
            Dispose_Game
            New_Game 4
            AutokKiirasa
        Case Else
            CleanALText
            AddAlText "Hiba!"
    End Select
End Sub

Private Sub AutokKiirasa()
    If Started Then
        Exit Sub
    End If

    Dim i As Byte
    CleanALText

    For i = LBound(Autok) To AutokSzama
        AddAlText "[" & i & ". Aut�] Sz�ne: " & Autok(i).GetColor()
    Next i
End Sub

Private Sub CleanALText()
    AutoListaText.Text = ""
End Sub

Private Sub AddAlText(Szoveg As String)
    AutoListaText.Text = AutoListaText.Text & Szoveg & vbCrLf
End Sub

