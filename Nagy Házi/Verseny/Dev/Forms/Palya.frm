VERSION 5.00
Begin VB.Form Palya 
   BackColor       =   &H8000000E&
   Caption         =   "Verseny Szimuláció"
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
      Caption         =   "Autók"
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
      Caption         =   "Kör: 0/0"
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
      Caption         =   "Játék"
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
         Caption         =   "Új játék"
         Shortcut        =   ^G
      End
      Begin VB.Menu Vegeredmeny_Mentese 
         Caption         =   "Végeredmény mentése"
      End
      Begin VB.Menu gamebar2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Kilpés"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "Beállítások"
      Begin VB.Menu Nyomvonal 
         Caption         =   "Nyomvonal"
         Shortcut        =   ^N
      End
      Begin VB.Menu Tokeletes_Korozes 
         Caption         =   "Tökéletes körözés"
         Shortcut        =   ^T
      End
      Begin VB.Menu settingbar 
         Caption         =   "-"
      End
      Begin VB.Menu GlobalSettings 
         Caption         =   "Beállítások"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Súgó"
      Begin VB.Menu About 
         Caption         =   "Névjegy"
      End
   End
End
Attribute VB_Name = "Palya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VersenyAdatok elnevezésõ listát frissíti.
Private WithEvents Timer_VersenyAdatok As VB.Timer
Attribute Timer_VersenyAdatok.VB_VarHelpID = -1
' AutoLista elnevezésõ listát frissíti.
Private WithEvents Timer_AutoLista As VB.Timer
Attribute Timer_AutoLista.VB_VarHelpID = -1
' Frissíti a körök számát. (Ha új kör van megváltoztatja a számlálót is.)
Private WithEvents Timer_Korok As VB.Timer
Attribute Timer_Korok.VB_VarHelpID = -1
' Tárolja éppen hányadik körnél tartunk.
Private Korok As Byte
' Jelzi hogy elindult-e már a játék vagy sem.
Private Started As Boolean
' Hány autó van kiválasztva. (terheléscsökkentés)
Private TempAutoLista As String
' Ha használva van a Stop gomb akkor lesz "true" az értéke.
Private Felfuggesztes As Boolean
' Tárolja hogy hányas számtól induljon az elsõ kör.
Private Const KezdokorErteke = 1
' Autók vonalának szélessége.
Private Const BorderWidth = 2
Private Const ex = 0.6
Private Const ey = -1

' Publikus változók.

' Visszadja publikusan a kezdõkör értékét.
Public Property Get GetKezdokorErteke() As Byte
    ' Érték beállítása.
    GetKezdokorErteke = KezdokorErteke
End Property

' Visszadja publikusan a körök számát. (Jelenlegi kör száma.)
Public Property Get GetKorokSzama() As Byte
    ' Érték beállítása.
    GetKorokSzama = Korok
End Property

' Publikus változók vége.

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Form_Load()
    ' Frissíti a virtuálisan létrehozott pályát.
    VirtualisPalya_Frissites

    ' Korok timer létrehozása
    Set Timer_Korok = Palya.Controls.Add("VB.Timer", "Timer_Korok", Palya)
    ' Érték beállítása. 40 millisec
    Timer_Korok.Interval = 40

    ' VersenyAdatok timer létrehozása
    Set Timer_VersenyAdatok = Palya.Controls.Add("VB.Timer", "Timer_VersenyAdatok", Palya)
    ' Érték beállítása. 500 millisec
    Timer_VersenyAdatok.Interval = 500

    ' AutoLista timer létrehozása
    Set Timer_AutoLista = Palya.Controls.Add("VB.Timer", "Timer_AutoLista", Palya)
    ' Érték beállítása. 100 millisec
    Timer_AutoLista.Interval = 500

    ' Nyomvonal megjelenésének beállítása
    Nyomvonal.Checked = Config.Globalis_Nyomvonal

    ' Tökéletes körözés beállítása
    Tokeletes_Korozes.Checked = Config.Globalis_TokeletesKorozes

    ' Alapértékek beállítása/takarítás.
    Clean
End Sub

' A form aktiválásakor lefutó vizsgálatok.
Private Sub Form_Activate()
    ' Megvizsgálja minden adat megfelelõ-e vagy sem. Ha nem le fog állni a program.
    Vizsgalat
End Sub

' Form megszünésekor bizonyos dolgok megsemisítésre kerülnek.
Private Sub Form_Terminate()
    ' Nullázás
    Set Timer_Korok = Nothing
    ' Nullázás
    Set Timer_VersenyAdatok = Nothing
    ' Nullázás
    Set Timer_AutoLista = Nothing
End Sub

' Form bezárása.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Minden form bezárása egyszerre. (Kilépés a programból.)
    Forms_Unload
End Sub

' NewGame menü gomb eseménye kattintás hatására.
Public Sub NewGame_Click()
    ' Megvizsgáljuk engedélyezve van-e az új játék indítása. Ha igen akkor töröljük a régit.
    If NewGameEnabled Then
        ' Játék törlése.
        Dispose_Game
        ' Alapértékek beállítása/takarítás.
        Clean
    End If

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    ' Azért fut le a ciklus hogy ellenörizzük minden autó befejezte-e a játékot.
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha egyik autó nem fejezte még be a játékot.
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            ' Kilépés a ciklusból.
            Exit For
        End If
    Next i

    ' Akkor fut le ha minden autó befejezte a játékot.
    If i = PalyaInfo.AutokSzama + 1 Then
        ' Megnyítja a figyelmeztetés ablakot jelezve hogy egy játék teljesen befejezödött.
        ' Így ha kivánja a felhasználó elmentheti a végeredményt.
        WarningNewGame.Show
    Else
        ' Játék törlése.
        Dispose_Game
        ' Alapértékek beállítása/takarítás.
        Clean
    End If
End Sub

' Új játék létrehozása.
' Az "ASzama" változó megfelel az autók számával. Azt tárolja hány autó lesz létrehozva a játékhoz.
Private Sub New_Game(ByVal ASzama As Byte)
    ' Megnézi fut-e már a játék. Ha igen akkor kilép az eljárásól.
    If Started Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Frissíti a virtuálisan létrehozott pályát.
    VirtualisPalya_Frissites

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    ' Létrehozunk és egyben újradimenzionálunk egy tomböt melynek neve "T".
    ' A tömb nagysága megfog egyezni az "Autok" tömb nagyságával.
    ReDim T(LBound(PalyaInfo.Autok) To UBound(PalyaInfo.Autok)) As String

    ' Értékadás a tömb elsõ elemének.
    T(1) = "piros"
    ' Értékadás a tömb második elemének.
    T(2) = "kék"
    ' Értékadás a tömb harmadik elemének.
    T(3) = "fekete"
    ' Értékadás a tömb negyedik elemének.
    T(4) = "zöld"

    For i = LBound(PalyaInfo.Autok) To ASzama
        ' Az autóra vonatkozó beállítások betöltése. Nem mindent töltünk be csak az alapokközül párat.
        PalyaInfo.Autok(i).Load i
        ' EX érték átadása.
        PalyaInfo.Autok(i).SetEX ex
        ' EY érték átadása.
        PalyaInfo.Autok(i).SetEY ey

        ' Akkor fut le ha a "KocsiVonalakSzama" tömb nagyobb vagy egyenlõ az autók számával.
        ' Vagy nagyobb nullánál.
        If PalyaInfo.KocsiVonalakSzama - 1 >= ASzama And PalyaInfo.KocsiVonalakSzama > 0 Then
            ' X0 koordináták átadása.
            PalyaInfo.Autok(i).SetX0 PalyaInfo.KocsiVonalTomb(i).X1
            ' Y0 koordináták átadása.
            PalyaInfo.Autok(i).SetY0 PalyaInfo.KocsiVonalTomb(i).Y1
        Else
            'Alapértelmezett X0 koordináták átadása.
            PalyaInfo.Autok(i).SetX0 1100
            'Alapértelmezett Y0 koordináták átadása.
            PalyaInfo.Autok(i).SetY0 5000
        End If

        ' Autó szinének átadása.
        PalyaInfo.Autok(i).SetColor T(i)
        ' Autó vonalainak vastagságának átadása.
        PalyaInfo.Autok(i).SetBorderWidth BorderWidth
        ' Autó megjelenítése.
        PalyaInfo.Autok(i).Show
    Next i

    ' Letároljuk hány autó van.
    PalyaInfo.AutokSzama = i - 1
End Sub

' Játék törlése.
Private Sub Dispose_Game()
    ' Frissíti a virtuálisan létrehozott pályát.
    VirtualisPalya_Frissites

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Törli az autó tulajdonásigait/beállításait.
        PalyaInfo.Autok(i).Dispose
        ' Törli az autótó.
        Set PalyaInfo.Autok(i) = Nothing
    Next i

    ' Az autók számát 0-ra állítja.
    PalyaInfo.AutokSzama = 0
End Sub

' Takarítja a futás közben felhalmozódott adatokat.
Private Sub Clean()
    ' Játék még csak most fog kezdõdni így az értéke "false" lesz.
    Started = False
    ' Végeredmény ablak bezárása.
    Unload VForm
    ' Játék törlése.
    Dispose_Game
    ' Az autók számát 0-ra állítja.
    PalyaInfo.AutokSzama = 0
    ' Felfüggesztés "false" értékre állítása.
    Felfuggesztes = False
    ' Új játék indításának lehetõségét "false"-ra állítjuk.
    NewGameEnabled = False
    ' Engedélyezzük a Timer_Korok idõzítöt.
    Timer_Korok.Enabled = True
    ' Engedélyezzük a Timer_AutoLista idõzítöt.
    Timer_AutoLista.Enabled = True
    ' Engedélyezzük az AutoLista combobox-ot.
    AutoLista.Enabled = True
    ' Beállítjuk hogy mitõl kezze el a körök számolását a rendszer.
    Korok = KezdokorErteke
    ' Megváltoztatjuk a körök számának kiírást.
    SetKorokSzama Korok

    ' Kezdõelem beállítása.
    VersenyAdatok.ListIndex = 0
    ' Kezdõelem beállítása.
    AutoLista.ListIndex = 0

    ' "i" segédváltozó a ciklushoz.
    Dim i As Integer
    For i = 0 To VersenyAdatok.ListCount
        ' Ha az "Autók sorrendje" elemel egyenlõ lesz az elem akkor eltároljuk az indexét.
        If "Autók sorrendje" = VersenyAdatok.List(i) Then
            ' Kezdõelem beállítása.
            VersenyAdatok.ListIndex = i
        End If
    Next i

    ' TempAutoLista takarítása.
    TempAutoLista = ""

    ' Tömb újradimenzionálása hogy nullázuk az elemeket.
    ReDim PalyaInfo.SorrendTomb(KezdokorErteke To Config.Globalis_KorokSzama) As Sorrend
End Sub

' GlobalSettings menü gomb eseménye kattintás hatására.
Private Sub GlobalSettings_Click()
    ' Megnézi fut-e már a játék. Ha igen akkor kilép az eljárásól.
    If Started Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Beállítások: Hiba!", "A játék már fut! Indíts új játékot ha szeretnél a beállíátsokon változtatni.", False
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Beállítások megjelenítése.
    SettingsForm.Show
End Sub

' Nyomvonal menü gomb eseménye kattintás hatására.
Private Sub Nyomvonal_Click()
    ' Ha igaz az érték akkor fut le.
    If Nyomvonal.Checked Then
        ' Mivel eddig igaz volt ezért hamisra állítjuk. Így kikapcsoljuk a pipát.
        Nyomvonal.Checked = False
    Else
        ' Mivel eddig hamis volt ezért igazra állítjuk. Így bekapcsoljuk a pipát.
        Nyomvonal.Checked = True
    End If

    ' Globális Nyomvonal változó frissítése.
    Config.Globalis_Nyomvonal = Nyomvonal.Checked
    ' Konfig fájl frissítése.
    Config.SetConfig
    ' Frissítési az autók nyomvonalának megjelenítését.
    SetAutokNyomvonal
End Sub

Public Sub SetAutokNyomvonal()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Szám szerint szelektál. Így mindig csak az adott autó nyomvonala változik meg.
        Select Case i
            Case 1
                ' Nyomvonal be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Elso_Nyomvonal)
            Case 2
                ' Nyomvonal be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Masodik_Nyomvonal)
            Case 3
                ' Nyomvonal be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Harmadik_Nyomvonal)
            Case 4
                ' Nyomvonal be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Negyedik_Nyomvonal)
        End Select
    Next i
End Sub

' Tokeletes_Korozes menü gomb eseménye kattintás hatására.
Private Sub Tokeletes_Korozes_Click()
    ' Ha igaz az érték akkor fut le.
    If Tokeletes_Korozes.Checked Then
        ' Mivel eddig igaz volt ezért hamisra állítjuk. Így kikapcsoljuk a pipát.
        Tokeletes_Korozes.Checked = False
    Else
        ' Mivel eddig hamis volt ezért igazra állítjuk. Így bekapcsoljuk a pipát.
        Tokeletes_Korozes.Checked = True
    End If

    ' Globális TokeletesKorozes változó frissítése.
    Config.Globalis_TokeletesKorozes = Tokeletes_Korozes.Checked
    ' Konfig fájl frissítése.
    Config.SetConfig
    ' Frissítési az autók tökéletes körözésének állapotát.
    SetAutokTokeletesKorozes
End Sub

Public Sub SetAutokTokeletesKorozes()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Szám szerint szelektál. Így mindig csak az adott autó tökéletes körözése változik meg.
        Select Case i
            Case 1
                ' Tökéletes körözés be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Elso_TokeletesKorozes)
            Case 2
                ' Tökéletes körözés be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Masodik_TokeletesKorozes)
            Case 3
                ' Tökéletes körözés be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Harmadik_TokeletesKorozes)
            Case 4
                ' Tökéletes körözés be illetve kikapcsolása.
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Negyedik_TokeletesKorozes)
        End Select
    Next i
End Sub

' Szektor nevek beállítása.
Public Sub SetSzektorNevek()
    ' Megnézi van-e szektor név. Ha nincs akkor kilép az eljárásól.
    If PalyaInfo.SzektorNevekSzama = 0 Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' "i" segédváltozó a ciklushoz.
    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorNevTomb) To PalyaInfo.SzektorNevekSzama - 1
        ' Beállítja a globális változó alapján a szektornév láthatóságát.
        PalyaInfo.SzektorNevTomb(i).Label.Visible = Config.Globalis_SzektorNevek
    Next i
End Sub

' Szektor nevek beállítása.
Public Sub SetSzektorVonalak()
    ' Megnézi van-e szektor vonal. Ha nincs akkor kilép az eljárásól.
    If PalyaInfo.SzektorVonalakSzama = 0 Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' "i" segédváltozó a ciklushoz.
    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
        ' Beállítja a globális változó alapján a szektor vonal láthatóságát.
        PalyaInfo.SzektorVonalTomb(i).Vonal.Visible = Config.Globalis_SzektorVonalak
    Next i

    ' Beállítja a globális változó alapján a start/célvonal láthatóságát.
    PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.Visible = Config.Globalis_StartCelVonal
End Sub

' Start menü gomb eseménye kattintás hatására.
Private Sub Start_Click()
    ' Megnézi van-e autó. Ha nincs akkor kilép az eljárásól.
    If PalyaInfo.AutokSzama = 0 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiba!", "Még nincsenek kiválasztva autók!", False
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Megnézi fut-e már a játék. Ha nem akkor átírja az álapotát "true"-ra.
    If Not Started Then
        ' Játék elindítotnak tekintése.
        Started = True
    End If

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha az autó már befejezte a játékot.
        If PalyaInfo.Autok(i).GetGameEnd Then
            ' Figyelmeztetés ablak megnyítása.
            WarningWindow "Hiba!", "A játék végetért! Nem indíthatod már el Start-tal! Indíts új játékot ha újat kezdenél.", False
            ' Kilépés az eljárásból.
            Exit Sub
        End If
    Next i

    ' Akkor fut le ha már fut a játék.
    If Felfuggesztes Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Start: Hiba!", "A játék már fut!", False
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Akkor fut le ha még nem fut a játék.
    If Not Felfuggesztes Then
        ' Felfüggesztés igazra állítása.
        Felfuggesztes = True
    End If

    ' Frissítési az autók nyomvonalának megjelenítését.
    SetAutokNyomvonal
    ' Frissítési az autók tökéletes körözésének állapotát.
    SetAutokTokeletesKorozes

    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Elindítja az autót.
        PalyaInfo.Autok(i).Start
    Next i
End Sub

' Stop menü gomb eseménye kattintás hatására.
Private Sub Stop_Click()
    If Not Felfuggesztes Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Stop: Hiba!", "A játék nem fut!", False
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Akkor fut le ha fut a játék.
    If Felfuggesztes Then
        ' Felfüggesztés hamisra állítása.
        Felfuggesztes = False
    End If

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Megállítja az autót.
        PalyaInfo.Autok(i).Stop_Kocsi
    Next i
End Sub

' Vegeredmeny_Mentese menü gomb eseménye kattintás hatására.
Private Sub Vegeredmeny_Mentese_Click()
    ' Elmenti a végeredményt.
    VegeredmenyMentese.Save
End Sub

' About menü gomb eseménye kattintás hatására.
Private Sub About_Click()
    ' Névjegy ablak megnyítása.
    AboutForm.Show
End Sub

' Exit menü gomb eseménye kattintás hatására.
Private Sub Exit_Click()
    ' Program bezárása.
    Forms_Unload
End Sub

' Minden formot bezárunk. Így teljesen leáll a program.
Private Sub Forms_Unload()
    ' Program vége.
    End
End Sub

' Kiírt körök felíratának megváltoztatása.
'A "KorSz" változó az aktuális kör számát tartalmazza.
Public Sub SetKorokSzama(ByVal KorSz As Byte)
    ' Átállítja a "0/0" értéket az aktuális körszámra és a maximális körszámra.
    KorKiiras.Caption = "Kör: " & KorSz & "/" & Config.Globalis_KorokSzama
End Sub

' Korok idõzitõ Timer eseménye.
Private Sub Timer_Korok_Timer()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha a Korok kisebb mint az autónak a jelenlegi kör száma.
        If Korok < PalyaInfo.Autok(i).GetKorokSzama Then
            ' Érték beállítása.
            Korok = PalyaInfo.Autok(i).GetKorokSzama

            ' Akkor fut le ha a Korok nagyobb mint a beállított maximális kör száma.
            If Korok > Config.Globalis_KorokSzama Then
                ' Végeredmény megjelenítése.
                VForm.Show
                ' Korok idõzitõ kikapcsolása.
                Timer_Korok.Enabled = False
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Jelengei körszám frissítése.
            SetKorokSzama Korok
        End If
    Next i
End Sub

' VersenyAdatok idõzitõ Timer eseménye.
Private Sub Timer_VersenyAdatok_Timer()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Single

    ' Kiválasztott elem alapján szelektáljuk melyik kiírás jelenjen meg.
    Select Case VersenyAdatok.List(VersenyAdatok.ListIndex)
        Case "Autók sorrendje"
            ' TextBox takarítása.
            CleanVAText

            ' Ha a játék még nem fut akkor fut le.
            If Not Started Then
                ' Hibaüzenet kiírása.
                NoStartedGameVAText
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Ideiglenes köröket tárol.
            Dim tempkor As Byte
            ' Ideiglenes autók számát tárolja.
            Dim tempautok As Byte

            ' Ha a Korok nagyobb mint a maximális körök száma akkor fut le.
            If Korok > Config.Globalis_KorokSzama Then
                ' Érték beállítása. Azért -1 mert a változó a játék végén +1-el nagyobbra lett megnövelve.
                tempkor = Korok - 1
            Else
                ' Érték beállítása.
                tempkor = Korok
            End If

            ' Nullázás.
            tempautok = 0

            ' Végtelenségig futó ciklus
            Do While True
                For ciklus = 3 To 1 Step -1
                    For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                        ' Akkor fut le ha nincs szin beállítva (nincs autó) és a van adat is.
                        If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                            ' Kilépés a ciklusból.
                            Exit For
                        ' Akkor fut le ha van adat és az ideiglenes autók száma kisebb vagy engyenlõ az AutokSzama-val.
                        ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                            ' Szöveg kiírása.
                            AddVAText i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin
                            ' Megnöveljük 1-el az ideiglenes autók számát.
                            tempautok = tempautok + 1
                        End If

                        ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
                        If tempautok = PalyaInfo.AutokSzama Then
                            ' Kilépés a ciklusból.
                            Exit For
                        End If
                    Next i

                    ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
                    If tempautok = PalyaInfo.AutokSzama Then
                        ' Kilépés a ciklusból.
                        Exit For
                    End If
                Next ciklus

                ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kilépés a ciklusból.
                    Exit Do
                End If

                ' Akkor fut le ha az ideiglenes körök száma nagyobb mind a kezdõkör értéke.
                If tempkor > KezdokorErteke Then
                    ' Az ideiglenes körök számát csökkentjük eggyel.
                    tempkor = tempkor - 1
                Else
                    ' Kilépés a ciklusból.
                    Exit Do
                End If
            Loop

            ' Akkor fut le ha az ideiglenes autók száma egyenlõ nullával.
            If tempautok = 0 Then
                ' Szöveg kiírása.
                AddVAText "Nincs még sorrend!"
            Else
                ' Szöveg kiírása.
                AddVAText ""
            End If

            ' Szöveg kiírása.
            AddVAText "A sorrend mindig a következõ szektornál frissül!"
        Case "Legjobb 1. szektor"
            SzektoridoKiiras 1
        Case "Legjobb 2. szektor"
            SzektoridoKiiras 2
        Case "Legjobb 3. szektor"
            SzektoridoKiiras 3
        Case "Legjobb köridõ"
            ' Tárolja a legjobb köridõt.
            Dim Szam As Single
            ' Tárolja az autó színét.
            Dim Szin As String
            ' Tárolja a legjobb köridõ körének számát.
            Dim lkor As Byte
            ' Tárolja az autó számát.
            Dim aszam As Byte

            ' TextBox takarítása.
            CleanVAText

            ' Akkor fut le ha a játék még nem fut.
            If Not Started Then
                ' Hibaüzenet kiírása.
                NoStartedGameVAText
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenlõ a kezdõkör számával.
            If Korok = KezdokorErteke Then
                ' Szöveg kiírása.
                AddVAText "Nincs még mért köridõ!"
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Érték beállítása.
            Szam = KezdoSzektorido
            For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                ' Akkor fut le ha a legjobb köridõ kisebb mint az eddig tárolt.
                If Szam > PalyaInfo.Autok(i).GetLegjobbKorido Then
                    ' Autó számának letárolása.
                    aszam = i
                    ' Autó szinének letárolása.
                    Szin = PalyaInfo.Autok(i).GetColor
                    ' Autó legjobb köridejének letárolása.
                    Szam = PalyaInfo.Autok(i).GetLegjobbKorido
                    ' Autó legjobb köridejéhez tartozó kör letárolása.
                    lkor = PalyaInfo.Autok(i).GetLegjobbKoridoSzama
                End If
            Next i

            ' Akkor fut le ha a legjobb köridõ egyenlõ az alapértelmezett kezdõszektoridõvel.
            If Szam = KezdoSzektorido Then
                ' Szöveg kiírása.
                AddVAText "Nincs még mért köridõ!"
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Szöveg kiírása.
            AddVAText "Legjobb kör ideje: " & Szam & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Elsõ szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Második szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " másodperc"
            ' Szöveg kiírása.
            AddVAText ""
            ' Szöveg kiírása.
            AddVAText "Az idõ a(z) " & lkor & ". körben került beállításra."
            ' Szöveg kiírása.
            AddVAText "A(z) idõt beállította a " & Szin & " szinû autó."
        Case "Elméleti legjobb köridõ"
            ' Tárolja a három legjobb szektoridõt.
            Dim T(1 To 3) As Single
            ' Tárolja a szektoridõkhöz tartozó autó szineket.
            Dim TSzin(1 To 3) As String
            ' Tárolja az elméleti legjobb köridõt.
            Dim ljekorido As Single
            ' TextBox takarítása.
            CleanVAText

            ' Akkor fut le ha a játék még nem fut.
            If Not Started Then
                ' Hibaüzenet kiírása.
                NoStartedGameVAText
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenlõ a kezdõkör számával.
            If Korok = KezdokorErteke Then
                ' Szöveg kiírása.
                AddVAText "Nincs még mért köridõ!"
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            For i = LBound(T) To UBound(T)
                ' Legjobb szektoridõ értékének beállítása.
                T(i) = LegjobbSzektorido(i, TSzin(i))
                ' Legjobb elméleti köridõhõz hozzáadja a szektoridõt.
                ljekorido = ljekorido + T(i)
            Next i

            ' Szöveg kiírása.
            AddVAText "Legjobb elméleti köriõ: " & ljekorido & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Elsõ szektor ideje: " & T(1) & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Autó színe: " & TSzin(1)
            ' Szöveg kiírása.
            AddVAText ""
            ' Szöveg kiírása.
            AddVAText "Második szektor ideje: " & T(2) & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Autó színe: " & TSzin(2)
            ' Szöveg kiírása.
            AddVAText ""
            ' Szöveg kiírása.
            AddVAText "Harmadik szektor ideje: " & T(3) & " másodperc"
            ' Szöveg kiírása.
            AddVAText "Autó színe: " & TSzin(3)
        Case "Pálya hossza"
            ' TextBox takarítása.
            CleanVAText

            ' Akkor fut le ha a játék még nem fut.
            If Not Started Then
                ' Hibaüzenet kiírása.
                NoStartedGameVAText
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Akkor fut le ha a Korok egyenlõ a kezdõkör számával.
            If Korok = KezdokorErteke Then
                ' Szöveg kiírása.
                AddVAText "Még nem sikerült megmérni a pálya hosszát!"
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            ' Szöveg kiírása.
            AddVAText "Egy kör hossza: " & PalyaInfo.Autok(LBound(PalyaInfo.Autok)).GetEgyKorHossza & " m"
            ' Szöveg kiírása.
            AddVAText "Ez az érték csak egy körübelüli érték mely a mozgásváltozással van összefüggésben."
            ' Szöveg kiírása.
            AddVAText PalyaHosszanakLepteke & " egység (egy lépés) felel meg " & PalyaHosszanakLepteke & " méternek."
        Case "1. Autó"
            EgyeniAutoKiirasok 1
        Case "2. Autó"
            EgyeniAutoKiirasok 2
        Case "3. Autó"
            EgyeniAutoKiirasok 3
        Case "4. Autó"
            EgyeniAutoKiirasok 4
        Case Else
            ' TextBox takarítása.
            CleanVAText
            ' Szöveg kiírása.
            AddVAText "Hiba!"
    End Select
End Sub

' Kiírja egyes autókról az információkat.
' Az "aszam" tárolja az adott autónak a számát.
Private Sub EgyeniAutoKiirasok(ByVal aszam As Byte)
    ' Tárolja a legjobb köridõt.
    Dim Szam As Single
    ' Tárolja az autó színét.
    Dim Szin As String
    ' Tárolja a legjobb köridõ körének számát.
    Dim lkor As Byte

    ' TextBox takarítása.
    CleanVAText

    ' Akkor fut le ha a játék még nem fut.
    If Not Started Then
        ' Hibaüzenet kiírása.
        NoStartedGameVAText
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Akkor fut le ha a megadott autószám nagyobb mint amennyi autó összesen el lett indítva.
    If aszam > PalyaInfo.AutokSzama Then
        ' Szöveg kiírása.
        AddVAText "Ez az autó nem versenyezik!"
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Akkor fut le ha a Korok egyenlõ a kezdõkör számával.
    If Korok = KezdokorErteke Then
        ' Szöveg kiírása.
        AddVAText "Nincs még mért köridõ!"
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Érték beállítása.
    Szam = KezdoSzektorido

    ' Akkor fut le ha a legjobb köridõ kisebb mint az eddig tárolt.
    If Szam > PalyaInfo.Autok(aszam).GetLegjobbKorido Then
        ' Autó szinének letárolása.
        Szin = PalyaInfo.Autok(aszam).GetColor
        ' Autó legjobb köridejének letárolása.
        Szam = PalyaInfo.Autok(aszam).GetLegjobbKorido
        ' Autó legjobb köridejéhez tartozó kör letárolása.
        lkor = PalyaInfo.Autok(aszam).GetLegjobbKoridoSzama
    End If

    ' Akkor fut le ha a legjobb köridõ egyenlõ az alapértelmezett kezdõszektoridõvel.
    If Szam = KezdoSzektorido Then
        ' Szöveg kiírása.
        AddVAText "Nincs még mért köridõ!"
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Szöveg kiírása.
    AddVAText "Legjobb eredmények:"
    ' Szöveg kiírása.
    AddVAText ""
    ' Szöveg kiírása.
    AddVAText "Kör ideje: " & Szam & " másodperc"
    ' Szöveg kiírása.
    AddVAText "Elsõ szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " másodperc"
    ' Szöveg kiírása.
    AddVAText "Második szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " másodperc"
    ' Szöveg kiírása.
    AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " másodperc"
    ' Szöveg kiírása.
    AddVAText ""
    ' Szöveg kiírása.
    AddVAText "Az idõ a(z) " & lkor & ". körben került beállításra."
End Sub

' Legjobb szektoridõt írja ki.
' Az "a" tárolja az szektor számát.  A 'Szin" tárolja az autó szinét.
Private Function LegjobbSzektorido(ByVal a As Integer, ByRef Szin As String) As Single
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    ' Tárolja a legjobb szektoridõt.
    Dim Szam As Single
    ' TextBox takarítása.
    CleanVAText
    ' Érték beállítása.
    Szam = KezdoSzektorido

    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha a legjobb köridõ kisebb mint az eddig tárolt.
        If Szam > PalyaInfo.Autok(i).GetLegjobbSzektoridok(a) Then
            ' Autó szinének letárolása.
            Szin = PalyaInfo.Autok(i).GetColor
            ' Autó legjobb szektoridejének letárolása.
            Szam = PalyaInfo.Autok(i).GetLegjobbSzektoridok(a)
        End If
    Next i

    ' Érték beállítása.
    LegjobbSzektorido = Szam
End Function

' Kiírja a szektor idejét.
' Az "a" tárolja a szektor számát.
Private Sub SzektoridoKiiras(ByVal a As Integer)
    Dim Szam As Single, Szin As String
    ' TextBox takarítása.
    CleanVAText

    ' Akkor fut le ha a játék még nem fut.
    If Not Started Then
        ' Hibaüzenet kiírása.
        NoStartedGameVAText
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Beállítja a legjobb szektoridõ idejét.
    Szam = LegjobbSzektorido(a, Szin)

    ' Akkor fut le ha a Szam értéke egyenlõ a kezdõ szektoridõvel.
    If Szam = KezdoSzektorido Then
        ' Szöveg kiírása.
        AddVAText "Nincs még mért szektoridõ!"
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Szöveg kiírása.
    AddVAText "Legjobb szektoridõ: " & Szam & " másodperc"
    ' Szöveg kiírása.
    AddVAText "Az idõt beállította a " & Szin & " szinû autó."
End Sub

' Még nem indult el a játék hibaüzenet kiírása.
Private Sub NoStartedGameVAText()
    ' Érték beállítása.
    VersenyAdatokText.Text = "Még nem indult el a játék!"
End Sub

' VersenyAdatok TextBox adatainak törlése.
Private Sub CleanVAText()
    ' Érték beállítása.
    VersenyAdatokText.Text = ""
End Sub

' VersenyAdatok TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddVAText(ByVal Szoveg As String)
    ' Érték beállítása.
    VersenyAdatokText.Text = VersenyAdatokText.Text & Szoveg & vbCrLf
End Sub

' AutoLista idõzitõ Timer eseménye.
Private Sub Timer_AutoLista_Timer()
    ' Akkor fut le ha a játék már fut.
    If Started Then
        ' Letiltja az AutoLista megváltoztathatóságát.
        AutoLista.Enabled = False
        ' AutoLista idõzitõ kikapcsolása.
        Timer_AutoLista.Enabled = False
    End If

    ' Akkor fut le ha az ideiglene autó lista egyenlõ a kiválasztottal.
    If TempAutoLista = AutoLista.List(AutoLista.ListIndex) Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' A kiválasztott alapján lefut amelyik kell.
    Select Case AutoLista.List(AutoLista.ListIndex)
        Case "Kettõ autó"
             ' Két autó létrehozása.
            UjAutokLetrehozasa 2
        Case "Három autó"
            ' Három autó létrehozása.
            UjAutokLetrehozasa 3
        Case "Négy autó"
            ' Négy autó létrehozása.
            UjAutokLetrehozasa 4
        Case Else
            ' TextBox takarítása.
            CleanALText
            ' Szöveg kiírása.
            AddALText "Hiba!"
    End Select
End Sub

' Új autók létrehozására szólgál.
' Az "UjAutokSzama" tárolja hány autó lesz létrehozva.
Public Sub UjAutokLetrehozasa(ByVal UjAutokSzama As Byte)
    ' Érték átadása.
    TempAutoLista = AutoLista.List(AutoLista.ListIndex)
    ' Játék törlése.
    Dispose_Game
    ' Új játék létrehozása.
    New_Game UjAutokSzama
    ' Autók kiírása.
    AutokKiirasa
End Sub

' Kiírja az autókat.
Private Sub AutokKiirasa()
    ' Megnézi fut-e már a játék. Ha igen akkor kilép az eljárásól.
    If Started Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' TextBox takarítása.
    CleanALText

    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Szöveg kiírása.
        AddALText "[" & i & ". Autó] Színe: " & PalyaInfo.Autok(i).GetColor()
    Next i
End Sub

' AutoListaText TextBox adatainak törlése.
Private Sub CleanALText()
    ' Érték beállítása.
    AutoListaText.Text = ""
End Sub

' AutoListaText TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddALText(ByVal Szoveg As String)
    ' Érték beállítása.
    AutoListaText.Text = AutoListaText.Text & Szoveg & vbCrLf
End Sub

' Virtuális pálya frissítése.
Public Sub VirtualisPalya_Frissites()
    ' Szektor nevek láthatóságának beállítása.
    SetSzektorNevek
    ' Szektor vonalak láthatóságának beállítása.
    SetSzektorVonalak

    ' Ha létezik a StartCelVonal akkor fut le.
    If Not PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        ' Beállítja a globális változó alapján a start/célvonal láthatóságát.
        PalyaInfo.StartCelVonalNev.Label.Visible = Config.Globalis_StartCelVonalNeve
    End If

    ' Akkor fut le ha a szektor vonalak szama nagyobb mint nulla.
    If PalyaInfo.SzektorVonalakSzama > 0 Then
        ' Beállítja a globális változó alapján a szektor vonal láthatóságát.
        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.Visible = Config.Globalis_StartCelVonal
    End If

    ' Változó létrehozása (POINTAPI).
    Dim vPt As POINTAPI

    ' Több tulajdonság kezelése.
    With VirtualisPalya
        ' Konvertálás a 0, 0 képernyõ kordinátára.
        ClientToScreen .hWnd, vPt
        ' "Container" kordináták konvertálása.
        ScreenToClient .Container.hWnd, vPt
        ' Eltolás PictureBox DC.
        SetViewportOrgEx .hDC, -vPt.X, -vPt.Y, vPt
        ' VB átfestés.
        SendMessage .Container.hWnd, WM_PAINT, .hDC, ByVal 0&
        ' Reset PictureBox DC.
        SetViewportOrgEx .hDC, vPt.X, vPt.Y, vPt

        ' Akkor fut le ha igaz az értéke.
        If .AutoRedraw Then
            ' Frissítés.
            .Refresh
        End If
    End With
End Sub
