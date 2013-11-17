VERSION 5.00
Begin VB.Form Palya 
   BackColor       =   &H8000000E&
   Caption         =   "Verseny Szimul�ci�"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   15465
   Begin VB.PictureBox VirtualisPalya 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   9855
      TabIndex        =   6
      Top             =   720
      Width           =   9855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Aut�k"
      Height          =   3495
      Left            =   10440
      TabIndex        =   2
      Top             =   840
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
         TabIndex        =   5
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox AutoLista 
         Height          =   315
         ItemData        =   "Palya.frx":0000
         Left            =   240
         List            =   "Palya.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Versenyadatok"
      Height          =   3615
      Left            =   10440
      TabIndex        =   0
      Top             =   4560
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
         TabIndex        =   4
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox VersenyAdatok 
         Height          =   315
         ItemData        =   "Palya.frx":0034
         Left            =   240
         List            =   "Palya.frx":0059
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Line MenuValasztoVonal 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   30000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label KorKiiras 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "K�r: 0/0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
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
      Begin VB.Menu Vegeredmeny_Mentese 
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
         Shortcut        =   ^N
      End
      Begin VB.Menu Tokeletes_Korozes 
         Caption         =   "T�k�letes k�r�z�s"
         Shortcut        =   ^T
      End
      Begin VB.Menu settingbar 
         Caption         =   "-"
      End
      Begin VB.Menu GlobalSettings 
         Caption         =   "Be�ll�t�sok"
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

' VersenyAdatok elnevez�s� list�t friss�ti.
Private WithEvents Timer_VersenyAdatok As VB.Timer
Attribute Timer_VersenyAdatok.VB_VarHelpID = -1
' AutoLista elnevez�s� list�t friss�ti.
Private WithEvents Timer_AutoLista As VB.Timer
Attribute Timer_AutoLista.VB_VarHelpID = -1
' Friss�ti a k�r�k sz�m�t. (Ha �j k�r van megv�ltoztatja a sz�ml�l�t is.)
Private WithEvents Timer_Korok As VB.Timer
Attribute Timer_Korok.VB_VarHelpID = -1
' T�rolja �ppen h�nyadik k�rn�l tartunk.
Private Korok As Byte
' Jelzi hogy elindult-e m�r a j�t�k vagy sem.
Private Started As Boolean
' H�ny aut� van kiv�lasztva. (terhel�scs�kkent�s)
Private TempAutoLista As String
' Ha haszn�lva van a Stop gomb akkor lesz "true" az �rt�ke.
Private Felfuggesztes As Boolean
' T�rolja hogy h�nyas sz�mt�l induljon az els� k�r.
Private Const KezdokorErteke = 1
' Aut�k vonal�nak sz�less�ge.
Private Const BorderWidth = 2
Private Const ex = 0.6
Private Const ey = -1

' Publikus v�ltoz�k.

' Visszadja publikusan a kezd�k�r �rt�k�t.
Public Property Get GetKezdokorErteke() As Byte
    ' �rt�k be�ll�t�sa.
    GetKezdokorErteke = KezdokorErteke
End Property

' Visszadja publikusan a k�r�k sz�m�t. (Jelenlegi k�r sz�ma.)
Public Property Get GetKorokSzama() As Byte
    ' �rt�k be�ll�t�sa.
    GetKorokSzama = Korok
End Property

' Publikus v�ltoz�k v�ge.

' Be�ll�tjuk a form l�trehoz�sakor az alap folyamatokat.
Private Sub Form_Load()
    ' Friss�ti a virtu�lisan l�trehozott p�ly�t.
    VirtualisPalya_Frissites

    ' Korok timer l�trehoz�sa
    Set Timer_Korok = Palya.Controls.Add("VB.Timer", "Timer_Korok", Palya)
    ' �rt�k be�ll�t�sa. 40 millisec
    Timer_Korok.Interval = 40

    ' VersenyAdatok timer l�trehoz�sa
    Set Timer_VersenyAdatok = Palya.Controls.Add("VB.Timer", "Timer_VersenyAdatok", Palya)
    ' �rt�k be�ll�t�sa. 500 millisec
    Timer_VersenyAdatok.Interval = 500

    ' AutoLista timer l�trehoz�sa
    Set Timer_AutoLista = Palya.Controls.Add("VB.Timer", "Timer_AutoLista", Palya)
    ' �rt�k be�ll�t�sa. 100 millisec
    Timer_AutoLista.Interval = 500

    ' Nyomvonal megjelen�s�nek be�ll�t�sa
    Nyomvonal.Checked = Config.Globalis_Nyomvonal

    ' T�k�letes k�r�z�s be�ll�t�sa
    Tokeletes_Korozes.Checked = Config.Globalis_TokeletesKorozes

    ' Alap�rt�kek be�ll�t�sa/takar�t�s.
    Clean
End Sub

' A form aktiv�l�sakor lefut� vizsg�latok.
Private Sub Form_Activate()
    ' Megvizsg�lja minden adat megfelel�-e vagy sem. Ha nem le fog �llni a program.
    Vizsgalat
End Sub

' Form megsz�n�sekor bizonyos dolgok megsemis�t�sre ker�lnek.
Private Sub Form_Terminate()
    ' Null�z�s
    Set Timer_Korok = Nothing
    ' Null�z�s
    Set Timer_VersenyAdatok = Nothing
    ' Null�z�s
    Set Timer_AutoLista = Nothing
End Sub

' Form bez�r�sa.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Minden form bez�r�sa egyszerre. (Kil�p�s a programb�l.)
    Forms_Unload
End Sub

' NewGame men� gomb esem�nye kattint�s hat�s�ra.
Public Sub NewGame_Click()
    ' Megvizsg�ljuk enged�lyezve van-e az �j j�t�k ind�t�sa. Ha igen akkor t�r�lj�k a r�git.
    If NewGameEnabled Then
        ' J�t�k t�rl�se.
        Dispose_Game
        ' Alap�rt�kek be�ll�t�sa/takar�t�s.
        Clean
    End If

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    ' Az�rt fut le a ciklus hogy ellen�rizz�k minden aut� befejezte-e a j�t�kot.
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha egyik aut� nem fejezte m�g be a j�t�kot.
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            ' Kil�p�s a ciklusb�l.
            Exit For
        End If
    Next i

    ' Akkor fut le ha minden aut� befejezte a j�t�kot.
    If i = PalyaInfo.AutokSzama + 1 Then
        ' Megny�tja a figyelmeztet�s ablakot jelezve hogy egy j�t�k teljesen befejez�d�tt.
        ' �gy ha kiv�nja a felhaszn�l� elmentheti a v�geredm�nyt.
        WarningNewGame.Show
    Else
        ' J�t�k t�rl�se.
        Dispose_Game
        ' Alap�rt�kek be�ll�t�sa/takar�t�s.
        Clean
    End If
End Sub

' �j j�t�k l�trehoz�sa.
' Az "ASzama" v�ltoz� megfelel az aut�k sz�m�val. Azt t�rolja h�ny aut� lesz l�trehozva a j�t�khoz.
Private Sub New_Game(ByVal ASzama As Byte)
    ' Megn�zi fut-e m�r a j�t�k. Ha igen akkor kil�p az elj�r�s�l.
    If Started Then
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Friss�ti a virtu�lisan l�trehozott p�ly�t.
    VirtualisPalya_Frissites

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    ' L�trehozunk �s egyben �jradimenzion�lunk egy tomb�t melynek neve "T".
    ' A t�mb nagys�ga megfog egyezni az "Autok" t�mb nagys�g�val.
    ReDim T(LBound(PalyaInfo.Autok) To UBound(PalyaInfo.Autok)) As String

    ' �rt�kad�s a t�mb els� elem�nek.
    T(1) = "piros"
    ' �rt�kad�s a t�mb m�sodik elem�nek.
    T(2) = "k�k"
    ' �rt�kad�s a t�mb harmadik elem�nek.
    T(3) = "fekete"
    ' �rt�kad�s a t�mb negyedik elem�nek.
    T(4) = "z�ld"

    For i = LBound(PalyaInfo.Autok) To ASzama
        ' Az aut�ra vonatkoz� be�ll�t�sok bet�lt�se. Nem mindent t�lt�nk be csak az alapokk�z�l p�rat.
        PalyaInfo.Autok(i).Load i
        ' EX �rt�k �tad�sa.
        PalyaInfo.Autok(i).SetEX ex
        ' EY �rt�k �tad�sa.
        PalyaInfo.Autok(i).SetEY ey

        ' Akkor fut le ha a "KocsiVonalakSzama" t�mb nagyobb vagy egyenl� az aut�k sz�m�val.
        ' Vagy nagyobb null�n�l.
        If PalyaInfo.KocsiVonalakSzama - 1 >= ASzama And PalyaInfo.KocsiVonalakSzama > 0 Then
            ' X0 koordin�t�k �tad�sa.
            PalyaInfo.Autok(i).SetX0 PalyaInfo.KocsiVonalTomb(i).X1
            ' Y0 koordin�t�k �tad�sa.
            PalyaInfo.Autok(i).SetY0 PalyaInfo.KocsiVonalTomb(i).Y1
        Else
            'Alap�rtelmezett X0 koordin�t�k �tad�sa.
            PalyaInfo.Autok(i).SetX0 1100
            'Alap�rtelmezett Y0 koordin�t�k �tad�sa.
            PalyaInfo.Autok(i).SetY0 5000
        End If

        ' Aut� szin�nek �tad�sa.
        PalyaInfo.Autok(i).SetColor T(i)
        ' Aut� vonalainak vastags�g�nak �tad�sa.
        PalyaInfo.Autok(i).SetBorderWidth BorderWidth
        ' Aut� megjelen�t�se.
        PalyaInfo.Autok(i).Show
    Next i

    ' Let�roljuk h�ny aut� van.
    PalyaInfo.AutokSzama = i - 1
End Sub

' J�t�k t�rl�se.
Private Sub Dispose_Game()
    ' Friss�ti a virtu�lisan l�trehozott p�ly�t.
    VirtualisPalya_Frissites

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' T�rli az aut� tulajdon�sigait/be�ll�t�sait.
        PalyaInfo.Autok(i).Dispose
        ' T�rli az aut�t�.
        Set PalyaInfo.Autok(i) = Nothing
    Next i

    ' Az aut�k sz�m�t 0-ra �ll�tja.
    PalyaInfo.AutokSzama = 0
End Sub

' Takar�tja a fut�s k�zben felhalmoz�dott adatokat.
Private Sub Clean()
    ' J�t�k m�g csak most fog kezd�dni �gy az �rt�ke "false" lesz.
    Started = False
    ' V�geredm�ny ablak bez�r�sa.
    Unload VForm
    ' J�t�k t�rl�se.
    Dispose_Game
    ' Az aut�k sz�m�t 0-ra �ll�tja.
    PalyaInfo.AutokSzama = 0
    ' Felf�ggeszt�s "false" �rt�kre �ll�t�sa.
    Felfuggesztes = False
    ' �j j�t�k ind�t�s�nak lehet�s�g�t "false"-ra �ll�tjuk.
    NewGameEnabled = False
    ' Enged�lyezz�k a Timer_Korok id�z�t�t.
    Timer_Korok.Enabled = True
    ' Enged�lyezz�k a Timer_AutoLista id�z�t�t.
    Timer_AutoLista.Enabled = True
    ' Enged�lyezz�k az AutoLista combobox-ot.
    AutoLista.Enabled = True
    ' Be�ll�tjuk hogy mit�l kezze el a k�r�k sz�mol�s�t a rendszer.
    Korok = KezdokorErteke
    ' Megv�ltoztatjuk a k�r�k sz�m�nak ki�r�st.
    SetKorokSzama Korok

    ' Kezd�elem be�ll�t�sa.
    VersenyAdatok.ListIndex = 0
    ' Kezd�elem be�ll�t�sa.
    AutoLista.ListIndex = 0

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Integer
    For i = 0 To VersenyAdatok.ListCount
        ' Ha az "Aut�k sorrendje" elemel egyenl� lesz az elem akkor elt�roljuk az index�t.
        If "Aut�k sorrendje" = VersenyAdatok.List(i) Then
            ' Kezd�elem be�ll�t�sa.
            VersenyAdatok.ListIndex = i
        End If
    Next i

    ' TempAutoLista takar�t�sa.
    TempAutoLista = ""

    ' T�mb �jradimenzion�l�sa hogy null�zuk az elemeket.
    ReDim PalyaInfo.SorrendTomb(KezdokorErteke To Config.Globalis_KorokSzama) As Sorrend
End Sub

' GlobalSettings men� gomb esem�nye kattint�s hat�s�ra.
Private Sub GlobalSettings_Click()
    ' Megn�zi fut-e m�r a j�t�k. Ha igen akkor kil�p az elj�r�s�l.
    If Started Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Be�ll�t�sok: Hiba!", "A j�t�k m�r fut! Ind�ts �j j�t�kot ha szeretn�l a be�ll��tsokon v�ltoztatni.", False
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Be�ll�t�sok megjelen�t�se.
    SettingsForm.Show
End Sub

' Nyomvonal men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Nyomvonal_Click()
    ' Ha igaz az �rt�k akkor fut le.
    If Nyomvonal.Checked Then
        ' Mivel eddig igaz volt ez�rt hamisra �ll�tjuk. �gy kikapcsoljuk a pip�t.
        Nyomvonal.Checked = False
    Else
        ' Mivel eddig hamis volt ez�rt igazra �ll�tjuk. �gy bekapcsoljuk a pip�t.
        Nyomvonal.Checked = True
    End If

    ' Glob�lis Nyomvonal v�ltoz� friss�t�se.
    Config.Globalis_Nyomvonal = Nyomvonal.Checked
    ' Konfig f�jl friss�t�se.
    Config.SetConfig
    ' Friss�t�si az aut�k nyomvonal�nak megjelen�t�s�t.
    SetAutokNyomvonal
End Sub

Public Sub SetAutokNyomvonal()
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Sz�m szerint szelekt�l. �gy mindig csak az adott aut� nyomvonala v�ltozik meg.
        Select Case i
            Case 1
                ' Nyomvonal be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Elso_Nyomvonal)
            Case 2
                ' Nyomvonal be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Masodik_Nyomvonal)
            Case 3
                ' Nyomvonal be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Harmadik_Nyomvonal)
            Case 4
                ' Nyomvonal be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Negyedik_Nyomvonal)
        End Select
    Next i
End Sub

' Tokeletes_Korozes men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Tokeletes_Korozes_Click()
    ' Ha igaz az �rt�k akkor fut le.
    If Tokeletes_Korozes.Checked Then
        ' Mivel eddig igaz volt ez�rt hamisra �ll�tjuk. �gy kikapcsoljuk a pip�t.
        Tokeletes_Korozes.Checked = False
    Else
        ' Mivel eddig hamis volt ez�rt igazra �ll�tjuk. �gy bekapcsoljuk a pip�t.
        Tokeletes_Korozes.Checked = True
    End If

    ' Glob�lis TokeletesKorozes v�ltoz� friss�t�se.
    Config.Globalis_TokeletesKorozes = Tokeletes_Korozes.Checked
    ' Konfig f�jl friss�t�se.
    Config.SetConfig
    ' Friss�t�si az aut�k t�k�letes k�r�z�s�nek �llapot�t.
    SetAutokTokeletesKorozes
End Sub

Public Sub SetAutokTokeletesKorozes()
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Sz�m szerint szelekt�l. �gy mindig csak az adott aut� t�k�letes k�r�z�se v�ltozik meg.
        Select Case i
            Case 1
                ' T�k�letes k�r�z�s be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Elso_TokeletesKorozes)
            Case 2
                ' T�k�letes k�r�z�s be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Masodik_TokeletesKorozes)
            Case 3
                ' T�k�letes k�r�z�s be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Harmadik_TokeletesKorozes)
            Case 4
                ' T�k�letes k�r�z�s be illetve kikapcsol�sa.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Negyedik_TokeletesKorozes)
        End Select
    Next i
End Sub

' Szektor nevek be�ll�t�sa.
Public Sub SetSzektorNevek()
    ' Megn�zi van-e szektor n�v. Ha nincs akkor kil�p az elj�r�s�l.
    If PalyaInfo.SzektorNevekSzama = 0 Then
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorNevTomb) To PalyaInfo.SzektorNevekSzama - 1
        ' Be�ll�tja a glob�lis v�ltoz� alapj�n a szektorn�v l�that�s�g�t.
        PalyaInfo.SzektorNevTomb(i).Label.Visible = Config.Globalis_SzektorNevek
    Next i
End Sub

' Szektor nevek be�ll�t�sa.
Public Sub SetSzektorVonalak()
    ' Megn�zi van-e szektor vonal. Ha nincs akkor kil�p az elj�r�s�l.
    If PalyaInfo.SzektorVonalakSzama = 0 Then
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
        ' Be�ll�tja a glob�lis v�ltoz� alapj�n a szektor vonal l�that�s�g�t.
        PalyaInfo.SzektorVonalTomb(i).Vonal.Visible = Config.Globalis_SzektorVonalak
    Next i

    ' Be�ll�tja a glob�lis v�ltoz� alapj�n a start/c�lvonal l�that�s�g�t.
    PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.Visible = Config.Globalis_StartCelVonal
End Sub

' Start men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Start_Click()
    ' Megn�zi van-e aut�. Ha nincs akkor kil�p az elj�r�s�l.
    If PalyaInfo.AutokSzama = 0 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hiba!", "M�g nincsenek kiv�lasztva aut�k!", False
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Megn�zi fut-e m�r a j�t�k. Ha nem akkor �t�rja az �lapot�t "true"-ra.
    If Not Started Then
        ' J�t�k elind�totnak tekint�se.
        Started = True
    End If

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha az aut� m�r befejezte a j�t�kot.
        If PalyaInfo.Autok(i).GetGameEnd Then
            ' Figyelmeztet�s ablak megny�t�sa.
            WarningWindow "Hiba!", "A j�t�k v�get�rt! Nem ind�thatod m�r el Start-tal! Ind�ts �j j�t�kot ha �jat kezden�l.", False
            ' Kil�p�s az elj�r�sb�l.
            Exit Sub
        End If
    Next i

    ' Akkor fut le ha m�r fut a j�t�k.
    If Felfuggesztes Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Start: Hiba!", "A j�t�k m�r fut!", False
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Akkor fut le ha m�g nem fut a j�t�k.
    If Not Felfuggesztes Then
        ' Felf�ggeszt�s igazra �ll�t�sa.
        Felfuggesztes = True
    End If

    ' Friss�t�si az aut�k nyomvonal�nak megjelen�t�s�t.
    SetAutokNyomvonal
    ' Friss�t�si az aut�k t�k�letes k�r�z�s�nek �llapot�t.
    SetAutokTokeletesKorozes

    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Elind�tja az aut�t.
        PalyaInfo.Autok(i).Start
    Next i
End Sub

' Stop men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Stop_Click()
    If Not Felfuggesztes Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Stop: Hiba!", "A j�t�k nem fut!", False
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Akkor fut le ha fut a j�t�k.
    If Felfuggesztes Then
        ' Felf�ggeszt�s hamisra �ll�t�sa.
        Felfuggesztes = False
    End If

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Meg�ll�tja az aut�t.
        PalyaInfo.Autok(i).Stop_Kocsi
    Next i
End Sub

' Vegeredmeny_Mentese men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Vegeredmeny_Mentese_Click()
    ' Elmenti a v�geredm�nyt.
    VegeredmenyMentese.Save
End Sub

' About men� gomb esem�nye kattint�s hat�s�ra.
Private Sub About_Click()
    ' N�vjegy ablak megny�t�sa.
    AboutForm.Show
End Sub

' Exit men� gomb esem�nye kattint�s hat�s�ra.
Private Sub Exit_Click()
    ' Program bez�r�sa.
    Forms_Unload
End Sub

' Minden formot bez�runk. �gy teljesen le�ll a program.
Private Sub Forms_Unload()
    ' Program v�ge.
    End
End Sub

' Ki�rt k�r�k fel�rat�nak megv�ltoztat�sa.
'A "KorSz" v�ltoz� az aktu�lis k�r sz�m�t tartalmazza.
Public Sub SetKorokSzama(ByVal KorSz As Byte)
    ' �t�ll�tja a "0/0" �rt�ket az aktu�lis k�rsz�mra �s a maxim�lis k�rsz�mra.
    KorKiiras.Caption = "K�r: " & KorSz & "/" & Config.Globalis_KorokSzama
End Sub

' Korok id�zit� Timer esem�nye.
Private Sub Timer_Korok_Timer()
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha a Korok kisebb mint az aut�nak a jelenlegi k�r sz�ma.
        If Korok < PalyaInfo.Autok(i).GetKorokSzama Then
            ' �rt�k be�ll�t�sa.
            Korok = PalyaInfo.Autok(i).GetKorokSzama

            ' Akkor fut le ha a Korok nagyobb mint a be�ll�tott maxim�lis k�r sz�ma.
            If Korok > Config.Globalis_KorokSzama Then
                ' V�geredm�ny megjelen�t�se.
                VForm.Show
                ' Korok id�zit� kikapcsol�sa.
                Timer_Korok.Enabled = False
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Jelengei k�rsz�m friss�t�se.
            SetKorokSzama Korok
        End If
    Next i
End Sub

' VersenyAdatok id�zit� Timer esem�nye.
Private Sub Timer_VersenyAdatok_Timer()
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    ' "ciklus" seg�dv�ltoz� a ciklushoz.
    Dim ciklus As Single

    ' Kiv�lasztott elem alapj�n szelekt�ljuk melyik ki�r�s jelenjen meg.
    Select Case VersenyAdatok.List(VersenyAdatok.ListIndex)
        Case "Aut�k sorrendje"
            ' TextBox takar�t�sa.
            CleanVAText

            ' Ha a j�t�k m�g nem fut akkor fut le.
            If Not Started Then
                ' Hiba�zenet ki�r�sa.
                NoStartedGameVAText
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Ideiglenes k�r�ket t�rol.
            Dim tempkor As Byte
            ' Ideiglenes aut�k sz�m�t t�rolja.
            Dim tempautok As Byte

            ' Ha a Korok nagyobb mint a maxim�lis k�r�k sz�ma akkor fut le.
            If Korok > Config.Globalis_KorokSzama Then
                ' �rt�k be�ll�t�sa. Az�rt -1 mert a v�ltoz� a j�t�k v�g�n +1-el nagyobbra lett megn�velve.
                tempkor = Korok - 1
            Else
                ' �rt�k be�ll�t�sa.
                tempkor = Korok
            End If

            ' Null�z�s.
            tempautok = 0

            ' V�gtelens�gig fut� ciklus
            Do While True
                For ciklus = 3 To 1 Step -1
                    For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                        ' Akkor fut le ha nincs szin be�ll�tva (nincs aut�) �s a van adat is.
                        If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                            ' Kil�p�s a ciklusb�l.
                            Exit For
                        ' Akkor fut le ha van adat �s az ideiglenes aut�k sz�ma kisebb vagy engyenl� az AutokSzama-val.
                        ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                            ' Sz�veg ki�r�sa.
                            AddVAText i & ". Aut�: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin
                            ' Megn�velj�k 1-el az ideiglenes aut�k sz�m�t.
                            tempautok = tempautok + 1
                        End If

                        ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
                        If tempautok = PalyaInfo.AutokSzama Then
                            ' Kil�p�s a ciklusb�l.
                            Exit For
                        End If
                    Next i

                    ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
                    If tempautok = PalyaInfo.AutokSzama Then
                        ' Kil�p�s a ciklusb�l.
                        Exit For
                    End If
                Next ciklus

                ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kil�p�s a ciklusb�l.
                    Exit Do
                End If

                ' Akkor fut le ha az ideiglenes k�r�k sz�ma nagyobb mind a kezd�k�r �rt�ke.
                If tempkor > KezdokorErteke Then
                    ' Az ideiglenes k�r�k sz�m�t cs�kkentj�k eggyel.
                    tempkor = tempkor - 1
                Else
                    ' Kil�p�s a ciklusb�l.
                    Exit Do
                End If
            Loop

            ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� null�val.
            If tempautok = 0 Then
                ' Sz�veg ki�r�sa.
                AddVAText "Nincs m�g sorrend!"
            Else
                ' Sz�veg ki�r�sa.
                AddVAText ""
            End If

            ' Sz�veg ki�r�sa.
            AddVAText "A sorrend mindig a k�vetkez� szektorn�l friss�l!"
        Case "Legjobb 1. szektor"
            SzektoridoKiiras 1
        Case "Legjobb 2. szektor"
            SzektoridoKiiras 2
        Case "Legjobb 3. szektor"
            SzektoridoKiiras 3
        Case "Legjobb k�rid�"
            ' T�rolja a legjobb k�rid�t.
            Dim Szam As Single
            ' T�rolja az aut� sz�n�t.
            Dim Szin As String
            ' T�rolja a legjobb k�rid� k�r�nek sz�m�t.
            Dim lkor As Byte
            ' T�rolja az aut� sz�m�t.
            Dim aszam As Byte

            ' TextBox takar�t�sa.
            CleanVAText

            ' Akkor fut le ha a j�t�k m�g nem fut.
            If Not Started Then
                ' Hiba�zenet ki�r�sa.
                NoStartedGameVAText
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenl� a kezd�k�r sz�m�val.
            If Korok = KezdokorErteke Then
                ' Sz�veg ki�r�sa.
                AddVAText "Nincs m�g m�rt k�rid�!"
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' �rt�k be�ll�t�sa.
            Szam = KezdoSzektorido
            For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                ' Akkor fut le ha a legjobb k�rid� kisebb mint az eddig t�rolt.
                If Szam > PalyaInfo.Autok(i).GetLegjobbKorido Then
                    ' Aut� sz�m�nak let�rol�sa.
                    aszam = i
                    ' Aut� szin�nek let�rol�sa.
                    Szin = PalyaInfo.Autok(i).GetColor
                    ' Aut� legjobb k�ridej�nek let�rol�sa.
                    Szam = PalyaInfo.Autok(i).GetLegjobbKorido
                    ' Aut� legjobb k�ridej�hez tartoz� k�r let�rol�sa.
                    lkor = PalyaInfo.Autok(i).GetLegjobbKoridoSzama
                End If
            Next i

            ' Akkor fut le ha a legjobb k�rid� egyenl� az alap�rtelmezett kezd�szektorid�vel.
            If Szam = KezdoSzektorido Then
                ' Sz�veg ki�r�sa.
                AddVAText "Nincs m�g m�rt k�rid�!"
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Sz�veg ki�r�sa.
            AddVAText "Legjobb k�r ideje: " & Szam & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Els� szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "M�sodik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText ""
            ' Sz�veg ki�r�sa.
            AddVAText "Az id� a(z) " & lkor & ". k�rben ker�lt be�ll�t�sra."
            ' Sz�veg ki�r�sa.
            AddVAText "A(z) id�t be�ll�totta a " & Szin & " szin� aut�."
        Case "Elm�leti legjobb k�rid�"
            ' T�rolja a h�rom legjobb szektorid�t.
            Dim T(1 To 3) As Single
            ' T�rolja a szektorid�kh�z tartoz� aut� szineket.
            Dim TSzin(1 To 3) As String
            ' T�rolja az elm�leti legjobb k�rid�t.
            Dim ljekorido As Single
            ' TextBox takar�t�sa.
            CleanVAText

            ' Akkor fut le ha a j�t�k m�g nem fut.
            If Not Started Then
                ' Hiba�zenet ki�r�sa.
                NoStartedGameVAText
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenl� a kezd�k�r sz�m�val.
            If Korok = KezdokorErteke Then
                ' Sz�veg ki�r�sa.
                AddVAText "Nincs m�g m�rt k�rid�!"
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            For i = LBound(T) To UBound(T)
                ' Legjobb szektorid� �rt�k�nek be�ll�t�sa.
                T(i) = LegjobbSzektorido(i, TSzin(i))
                ' Legjobb elm�leti k�rid�h�z hozz�adja a szektorid�t.
                ljekorido = ljekorido + T(i)
            Next i

            ' Sz�veg ki�r�sa.
            AddVAText "Legjobb elm�leti k�ri�: " & ljekorido & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Els� szektor ideje: " & T(1) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Aut� sz�ne: " & TSzin(1)
            ' Sz�veg ki�r�sa.
            AddVAText ""
            ' Sz�veg ki�r�sa.
            AddVAText "M�sodik szektor ideje: " & T(2) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Aut� sz�ne: " & TSzin(2)
            ' Sz�veg ki�r�sa.
            AddVAText ""
            ' Sz�veg ki�r�sa.
            AddVAText "Harmadik szektor ideje: " & T(3) & " m�sodperc"
            ' Sz�veg ki�r�sa.
            AddVAText "Aut� sz�ne: " & TSzin(3)
        Case "P�lya hossza"
            ' TextBox takar�t�sa.
            CleanVAText

            ' Akkor fut le ha a j�t�k m�g nem fut.
            If Not Started Then
                ' Hiba�zenet ki�r�sa.
                NoStartedGameVAText
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenl� a kezd�k�r sz�m�val.
            If Korok = KezdokorErteke Then
                ' Sz�veg ki�r�sa.
                AddVAText "M�g nem siker�lt megm�rni a p�lya hossz�t!"
                ' Kil�p�s az elj�r�sb�l.
                Exit Sub
            End If

            ' Sz�veg ki�r�sa.
            AddVAText "Egy k�r hossza: " & PalyaInfo.Autok(LBound(PalyaInfo.Autok)).GetEgyKorHossza & " m"
            ' Sz�veg ki�r�sa.
            AddVAText "Ez az �rt�k csak egy k�r�bel�li �rt�k mely a mozg�sv�ltoz�ssal van �sszef�gg�sben."
            ' Sz�veg ki�r�sa.
            AddVAText PalyaHosszanakLepteke & " egys�g (egy l�p�s) felel meg " & PalyaHosszanakLepteke & " m�ternek."
        Case "1. Aut�"
            EgyeniAutoKiirasok 1
        Case "2. Aut�"
            EgyeniAutoKiirasok 2
        Case "3. Aut�"
            EgyeniAutoKiirasok 3
        Case "4. Aut�"
            EgyeniAutoKiirasok 4
        Case Else
            ' TextBox takar�t�sa.
            CleanVAText
            ' Sz�veg ki�r�sa.
            AddVAText "Hiba!"
    End Select
End Sub

' Ki�rja egyes aut�kr�l az inform�ci�kat.
' Az "aszam" t�rolja az adott aut�nak a sz�m�t.
Private Sub EgyeniAutoKiirasok(ByVal aszam As Byte)
    ' T�rolja a legjobb k�rid�t.
    Dim Szam As Single
    ' T�rolja az aut� sz�n�t.
    Dim Szin As String
    ' T�rolja a legjobb k�rid� k�r�nek sz�m�t.
    Dim lkor As Byte

    ' TextBox takar�t�sa.
    CleanVAText

    ' Akkor fut le ha a j�t�k m�g nem fut.
    If Not Started Then
        ' Hiba�zenet ki�r�sa.
        NoStartedGameVAText
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Akkor fut le ha a megadott aut�sz�m nagyobb mint amennyi aut� �sszesen el lett ind�tva.
    If aszam > PalyaInfo.AutokSzama Then
        ' Sz�veg ki�r�sa.
        AddVAText "Ez az aut� nem versenyezik!"
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Akkor fut le ha a Korok egyenl� a kezd�k�r sz�m�val.
    If Korok = KezdokorErteke Then
        ' Sz�veg ki�r�sa.
        AddVAText "Nincs m�g m�rt k�rid�!"
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' �rt�k be�ll�t�sa.
    Szam = KezdoSzektorido

    ' Akkor fut le ha a legjobb k�rid� kisebb mint az eddig t�rolt.
    If Szam > PalyaInfo.Autok(aszam).GetLegjobbKorido Then
        ' Aut� szin�nek let�rol�sa.
        Szin = PalyaInfo.Autok(aszam).GetColor
        ' Aut� legjobb k�ridej�nek let�rol�sa.
        Szam = PalyaInfo.Autok(aszam).GetLegjobbKorido
        ' Aut� legjobb k�ridej�hez tartoz� k�r let�rol�sa.
        lkor = PalyaInfo.Autok(aszam).GetLegjobbKoridoSzama
    End If

    ' Akkor fut le ha a legjobb k�rid� egyenl� az alap�rtelmezett kezd�szektorid�vel.
    If Szam = KezdoSzektorido Then
        ' Sz�veg ki�r�sa.
        AddVAText "Nincs m�g m�rt k�rid�!"
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Sz�veg ki�r�sa.
    AddVAText "Legjobb eredm�nyek:"
    ' Sz�veg ki�r�sa.
    AddVAText ""
    ' Sz�veg ki�r�sa.
    AddVAText "K�r ideje: " & Szam & " m�sodperc"
    ' Sz�veg ki�r�sa.
    AddVAText "Els� szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " m�sodperc"
    ' Sz�veg ki�r�sa.
    AddVAText "M�sodik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " m�sodperc"
    ' Sz�veg ki�r�sa.
    AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " m�sodperc"
    ' Sz�veg ki�r�sa.
    AddVAText ""
    ' Sz�veg ki�r�sa.
    AddVAText "Az id� a(z) " & lkor & ". k�rben ker�lt be�ll�t�sra."
End Sub

' Legjobb szektorid�t �rja ki.
' Az "a" t�rolja az szektor sz�m�t.  A 'Szin" t�rolja az aut� szin�t.
Private Function LegjobbSzektorido(ByVal a As Integer, ByRef Szin As String) As Single
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    ' T�rolja a legjobb szektorid�t.
    Dim Szam As Single
    ' TextBox takar�t�sa.
    CleanVAText
    ' �rt�k be�ll�t�sa.
    Szam = KezdoSzektorido

    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha a legjobb k�rid� kisebb mint az eddig t�rolt.
        If Szam > PalyaInfo.Autok(i).GetLegjobbSzektoridok(a) Then
            ' Aut� szin�nek let�rol�sa.
            Szin = PalyaInfo.Autok(i).GetColor
            ' Aut� legjobb szektoridej�nek let�rol�sa.
            Szam = PalyaInfo.Autok(i).GetLegjobbSzektoridok(a)
        End If
    Next i

    ' �rt�k be�ll�t�sa.
    LegjobbSzektorido = Szam
End Function

' Ki�rja a szektor idej�t.
' Az "a" t�rolja a szektor sz�m�t.
Private Sub SzektoridoKiiras(ByVal a As Integer)
    Dim Szam As Single, Szin As String
    ' TextBox takar�t�sa.
    CleanVAText

    ' Akkor fut le ha a j�t�k m�g nem fut.
    If Not Started Then
        ' Hiba�zenet ki�r�sa.
        NoStartedGameVAText
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Be�ll�tja a legjobb szektorid� idej�t.
    Szam = LegjobbSzektorido(a, Szin)

    ' Akkor fut le ha a Szam �rt�ke egyenl� a kezd� szektorid�vel.
    If Szam = KezdoSzektorido Then
        ' Sz�veg ki�r�sa.
        AddVAText "Nincs m�g m�rt szektorid�!"
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' Sz�veg ki�r�sa.
    AddVAText "Legjobb szektorid�: " & Szam & " m�sodperc"
    ' Sz�veg ki�r�sa.
    AddVAText "Az id�t be�ll�totta a " & Szin & " szin� aut�."
End Sub

' M�g nem indult el a j�t�k hiba�zenet ki�r�sa.
Private Sub NoStartedGameVAText()
    ' �rt�k be�ll�t�sa.
    VersenyAdatokText.Text = "M�g nem indult el a j�t�k!"
End Sub

' VersenyAdatok TextBox adatainak t�rl�se.
Private Sub CleanVAText()
    ' �rt�k be�ll�t�sa.
    VersenyAdatokText.Text = ""
End Sub

' VersenyAdatok TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddVAText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    VersenyAdatokText.Text = VersenyAdatokText.Text & Szoveg & vbCrLf
End Sub

' AutoLista id�zit� Timer esem�nye.
Private Sub Timer_AutoLista_Timer()
    ' Akkor fut le ha a j�t�k m�r fut.
    If Started Then
        ' Letiltja az AutoLista megv�ltoztathat�s�g�t.
        AutoLista.Enabled = False
        ' AutoLista id�zit� kikapcsol�sa.
        Timer_AutoLista.Enabled = False
    End If

    ' Akkor fut le ha az ideiglene aut� lista egyenl� a kiv�lasztottal.
    If TempAutoLista = AutoLista.List(AutoLista.ListIndex) Then
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' A kiv�lasztott alapj�n lefut amelyik kell.
    Select Case AutoLista.List(AutoLista.ListIndex)
        Case "Kett� aut�"
             ' K�t aut� l�trehoz�sa.
            UjAutokLetrehozasa 2
        Case "H�rom aut�"
            ' H�rom aut� l�trehoz�sa.
            UjAutokLetrehozasa 3
        Case "N�gy aut�"
            ' N�gy aut� l�trehoz�sa.
            UjAutokLetrehozasa 4
        Case Else
            ' TextBox takar�t�sa.
            CleanALText
            ' Sz�veg ki�r�sa.
            AddALText "Hiba!"
    End Select
End Sub

' �j aut�k l�trehoz�s�ra sz�lg�l.
' Az "UjAutokSzama" t�rolja h�ny aut� lesz l�trehozva.
Public Sub UjAutokLetrehozasa(ByVal UjAutokSzama As Byte)
    ' �rt�k �tad�sa.
    TempAutoLista = AutoLista.List(AutoLista.ListIndex)
    ' J�t�k t�rl�se.
    Dispose_Game
    ' �j j�t�k l�trehoz�sa.
    New_Game UjAutokSzama
    ' Aut�k ki�r�sa.
    AutokKiirasa
End Sub

' Ki�rja az aut�kat.
Private Sub AutokKiirasa()
    ' Megn�zi fut-e m�r a j�t�k. Ha igen akkor kil�p az elj�r�s�l.
    If Started Then
        ' Kil�p�s az elj�r�sb�l.
        Exit Sub
    End If

    ' TextBox takar�t�sa.
    CleanALText

    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Sz�veg ki�r�sa.
        AddALText "[" & i & ". Aut�] Sz�ne: " & PalyaInfo.Autok(i).GetColor()
    Next i
End Sub

' AutoListaText TextBox adatainak t�rl�se.
Private Sub CleanALText()
    ' �rt�k be�ll�t�sa.
    AutoListaText.Text = ""
End Sub

' AutoListaText TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddALText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    AutoListaText.Text = AutoListaText.Text & Szoveg & vbCrLf
End Sub

' Virtu�lis p�lya friss�t�se.
Public Sub VirtualisPalya_Frissites()
    ' Szektor nevek l�that�s�g�nak be�ll�t�sa.
    SetSzektorNevek
    ' Szektor vonalak l�that�s�g�nak be�ll�t�sa.
    SetSzektorVonalak

    ' Ha l�tezik a StartCelVonal akkor fut le.
    If Not PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        ' Be�ll�tja a glob�lis v�ltoz� alapj�n a start/c�lvonal l�that�s�g�t.
        PalyaInfo.StartCelVonalNev.Label.Visible = Config.Globalis_StartCelVonalNeve
    End If

    ' Akkor fut le ha a szektor vonalak szama nagyobb mint nulla.
    If PalyaInfo.SzektorVonalakSzama > 0 Then
        ' Be�ll�tja a glob�lis v�ltoz� alapj�n a szektor vonal l�that�s�g�t.
        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.Visible = Config.Globalis_StartCelVonal
    End If

    ' V�ltoz� l�trehoz�sa (POINTAPI).
    Dim vPt As POINTAPI

    ' T�bb tulajdons�g kezel�se.
    With VirtualisPalya
        ' Konvert�l�s a 0, 0 k�perny� kordin�t�ra.
        ClientToScreen .hWnd, vPt
        ' "Container" kordin�t�k konvert�l�sa.
        ScreenToClient .Container.hWnd, vPt
        ' Eltol�s PictureBox DC.
        SetViewportOrgEx .hDC, -vPt.X, -vPt.Y, vPt
        ' VB �tfest�s.
        SendMessage .Container.hWnd, WM_PAINT, .hDC, ByVal 0&
        ' Reset PictureBox DC.
        SetViewportOrgEx .hDC, vPt.X, vPt.Y, vPt

        ' Akkor fut le ha igaz az �rt�ke.
        If .AutoRedraw Then
            ' Friss�t�s.
            .Refresh
        End If
    End With
End Sub
