VERSION 5.00
Begin VB.Form Palya 
   BackColor       =   &H8000000E&
   Caption         =   "Verseny Szimuláció"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   15465
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox HamisPalya 
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
Private WithEvents Timer_VersenyAdatok As VB.Timer  ' VersenyAdatok elnevezésõ listát frissíti.
Attribute Timer_VersenyAdatok.VB_VarHelpID = -1
Private WithEvents Timer_AutoLista As VB.Timer      ' AutoLista elnevezésõ listát frissíti.
Attribute Timer_AutoLista.VB_VarHelpID = -1
Private WithEvents Timer_Korok As VB.Timer          ' Frissíti a körök számát. (Ha új kör van megváltoztatja a számlálót is.)
Attribute Timer_Korok.VB_VarHelpID = -1
Private AutokSzama As Byte                          ' Versenypályán lévõ autók számát tárolja.
Private Korok As Byte                               ' Tárolja éppen hányadik körnél tartunk.
Private Started As Boolean                          ' Jelzi hogy elindult-e már a játék vagy sem.
Public TempAutoLista As String                      ' Hány autó van kiválasztva. (terheléscsökkentés)
Private Felfuggesztes As Boolean                    ' Ha használva van a Stop gomb akkor lesz "true" az értéke.
Private Const KezdokorErteke = 1                    ' Tárolja hogy hánytól induljon az elsõ kör.
Private Const BorderWidth = 2                       ' Autók vonalának szélessége.
Private Const PalyaHosszanakLepteke = 5             ' 5 m-t jelent. Ez azt jelenti hogy egy elmozdulással az autó 10 métert tesz meg.
Private Const ex = 0.6
Private Const ey = -1

' Publikus változók.
Public Property Get GetPalyaHosszanakLepteke() As Byte
    GetPalyaHosszanakLepteke = PalyaHosszanakLepteke
End Property

Public Property Get GetKezdokorErteke() As Byte
    GetKezdokorErteke = KezdokorErteke
End Property

Public Property Get GetKorokSzama() As Byte
    GetKorokSzama = Korok
End Property

Public Property Get GetAutokSzama() As Byte
    GetAutokSzama = AutokSzama
End Property

' Publikus változók vége.

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Form_Load()
    HamisPalya_Frissites

    ' Korok timer létrehozása
    Set Timer_Korok = Palya.Controls.Add("VB.Timer", "Timer_Korok", Palya)
    Timer_Korok.Interval = 40          ' Érték beállítása. 40 millisec

    ' VersenyAdatok timer létrehozása
    Set Timer_VersenyAdatok = Palya.Controls.Add("VB.Timer", "Timer_VersenyAdatok", Palya)
    Timer_VersenyAdatok.Interval = 500 ' Érték beállítása. 500 millisec

    ' AutoLista timer létrehozása
    Set Timer_AutoLista = Palya.Controls.Add("VB.Timer", "Timer_AutoLista", Palya)
    Timer_AutoLista.Interval = 500     ' Érték beállítása. 100 millisec

    ' Nyomvonal megjelenésének beállítása
    Nyomvonal.Checked = Config.Globalis_Nyomvonal

    ' Tökéletes körözés beállítása
    Tokeletes_Korozes.Checked = Config.Globalis_TokeletesKorozes

    ' Alapértékek beállítása.
    Clean
End Sub

' Form megszünésekor bizonyos dolgok megsemisítésre kerülnek.
Private Sub Form_Terminate()
    Set Timer_Korok = Nothing         ' Nullázás
    Set Timer_VersenyAdatok = Nothing ' Nullázás
    Set Timer_AutoLista = Nothing     ' Nullázás
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Forms_Unload
End Sub

Private Sub New_Game(ASzama As Byte)
    If Started Then
        Exit Sub
    End If

    HamisPalya_Frissites

    Dim i As Byte
    i = 1

    Dim T(1 To 4) As String
    T(1) = "piros"
    T(2) = "kék"
    T(3) = "fekete"
    T(4) = "zöld"

    For i = LBound(PalyaInfo.Autok) To ASzama
        PalyaInfo.Autok(i).Load i ' Betöltjük újként a vonalat
        PalyaInfo.Autok(i).SetEX ex
        PalyaInfo.Autok(i).SetEY ey

        If PalyaInfo.KocsiVonalakSzama - 1 >= ASzama And PalyaInfo.KocsiVonalakSzama >= 1 Then
            PalyaInfo.Autok(i).SetX0 PalyaInfo.KocsiVonalTomb(i).X1
            PalyaInfo.Autok(i).SetY0 PalyaInfo.KocsiVonalTomb(i).Y1
        Else
            PalyaInfo.Autok(i).SetX0 1100
            PalyaInfo.Autok(i).SetY0 5000
        End If

        PalyaInfo.Autok(i).SetColor T(i) ' Ha kell színezés csak akkor.
        PalyaInfo.Autok(i).SetBorderWidth BorderWidth
        PalyaInfo.Autok(i).Show
    Next i

    AutokSzama = i - 1
End Sub

Private Sub Dispose_Game()
    If Started Then
        Exit Sub
    End If

    HamisPalya_Frissites

    Dim i As Byte
    i = 1

    For i = 1 To AutokSzama
        PalyaInfo.Autok(i).Dispose
        Set PalyaInfo.Autok(i) = Nothing
    Next i

    AutokSzama = 0
End Sub

Private Sub GlobalSettings_Click()
    If Started Then
        WarningWindow "Beállítások: Hiba!", "A játék már fut! Indíts új játékot ha szeretnél a beállíátsokon változtatni."
        Exit Sub
    End If

    SettingsForm.Show
End Sub

Public Sub NewGame_Click()
    If NewGameEnabled Then
        Dispose_Game
        Clean
    End If

    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            Exit For
        End If
    Next i

    If i = AutokSzama + 1 Then
        WarningNewGame.Show
    Else
        Dispose_Game
        Clean
    End If
End Sub

Private Sub Clean()
    Started = False                 ' Játék még csak most fog kezdõdni így az értéke "false" lesz.
    Unload VForm                    ' Végeredmény ablak bezárása.
    Dispose_Game                    ' Régi autók megsemmisítése.
    AutokSzama = 0                  ' Autók számát nullára állítjuk.
    Felfuggesztes = False           ' Felfüggesztés "false" értékre állítása.
    NewGameEnabled = False          ' Új játék indításának lehetõségét "false"-ra állítjuk.
    Timer_Korok.Enabled = True      ' Engedélyezzük a Korok timer-t.
    Timer_AutoLista.Enabled = True  ' Engedélyezzük az AutoLista timer-t.
    AutoLista.Enabled = True        ' Engedélyezzük az AutoLista combobox-ot.
    Korok = KezdokorErteke          ' Beállítjuk hogy mitõl kezze el a körök számolását a rendszer.
    SetKorokSzama Korok             ' Megváltoztatjuk a körök számának kiírást.

    VersenyAdatok.ListIndex = 0     ' Kezdõelem beállítása.
    AutoLista.ListIndex = 0         ' Kezdõelem beállítása.
    TempAutoLista = ""              ' Takarítás.

    ' Tömb újradimenzionálása a nullázás érdekében.
    ReDim PalyaInfo.SorrendTomb(KezdokorErteke To Config.Globalis_KorokSzama) As Sorrend
End Sub

Private Sub Nyomvonal_Click()
    If Nyomvonal.Checked Then
        Nyomvonal.Checked = False
    Else
        Nyomvonal.Checked = True
    End If

    Config.Globalis_Nyomvonal = Nyomvonal.Checked
    Config.SetConfig
    SetAutokNyomvonal
End Sub

Public Sub SetAutokNyomvonal()
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        Select Case i
            Case 1
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Elso_Nyomvonal)
            Case 2
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Masodik_Nyomvonal)
            Case 3
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Harmadik_Nyomvonal)
            Case 4
                PalyaInfo.Autok(i).SetNyomvonal (Config.Globalis_Nyomvonal And Config.Autok_Negyedik_Nyomvonal)
        End Select
    Next i
End Sub

Private Sub Tokeletes_Korozes_Click()
    If Tokeletes_Korozes.Checked Then
        Tokeletes_Korozes.Checked = False
    Else
        Tokeletes_Korozes.Checked = True
    End If

    Config.Globalis_TokeletesKorozes = Tokeletes_Korozes.Checked
    Config.SetConfig
    SetAutokTokeletesKorozes
End Sub

Public Sub SetAutokTokeletesKorozes()
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        Select Case i
            Case 1
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Elso_TokeletesKorozes)
            Case 2
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Masodik_TokeletesKorozes)
            Case 3
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Harmadik_TokeletesKorozes)
            Case 4
                PalyaInfo.Autok(i).SetPontosabbKor (Config.Globalis_TokeletesKorozes And Config.Autok_Negyedik_TokeletesKorozes)
        End Select
    Next i
End Sub

Public Sub SetSzektorNevek()
    If PalyaInfo.SzektorNevekSzama = 0 Then
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorNevTomb) To PalyaInfo.SzektorNevekSzama - 1
        PalyaInfo.SzektorNevTomb(i).Label.visible = Config.Globalis_SzektorNevek
    Next i
End Sub

Public Sub SetSzektorVonalak()
    If PalyaInfo.SzektorVonalakSzama = 0 Then
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
        PalyaInfo.SzektorVonalTomb(i).Vonal.visible = Config.Globalis_SzektorVonalak
    Next i

    PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.visible = Config.Globalis_StartCelVonal
End Sub

Private Sub Start_Click()
    If AutokSzama = 0 Then
        WarningWindow "Hiba!", "Még nincsenek kiválasztva autók!"
        Exit Sub
    End If

    If Not Started Then
        Started = True
    End If

    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        If PalyaInfo.Autok(i).GetGameEnd Then
            WarningWindow "Hiba!", "A játék végetért! Nem indíthatod már el Start-tal! Indíts új játékot ha újat kezdenél."
            Exit Sub
        End If
    Next i

    If Felfuggesztes Then
        WarningWindow "Start: Hiba!", "A játék már fut!"
        Exit Sub
    End If

    If Not Felfuggesztes Then
        Felfuggesztes = True
    End If

    SetAutokNyomvonal ' Frissítés hogy ne a hibás adat maradjon benne ha nem lett megváltoztatva.
    SetAutokTokeletesKorozes ' Frissítés hogy ne a hibás adat maradjon benne ha nem lett megváltoztatva.

    For i = LBound(PalyaInfo.Autok) To AutokSzama
        PalyaInfo.Autok(i).Start
    Next i
End Sub

Private Sub Stop_Click()
    If Not Felfuggesztes Then
        WarningWindow "Stop: Hiba!", "A játék nem fut!"
        Exit Sub
    End If

    If Felfuggesztes Then
        Felfuggesztes = False
    End If

    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        PalyaInfo.Autok(i).Stop_Kocsi
    Next i
End Sub

Private Sub Vegeredmeny_Mentese_Click()
    VegeredmenyMentese.Save
End Sub

' Névjegy ablak megnyítása.
Private Sub About_Click()
    AboutForm.Show
End Sub

' Program bezárása.
Private Sub Exit_Click()
    Forms_Unload
End Sub

' Minden formot bezárunk hogy nehogy valamelyik is nyítva maradjon.
Private Sub Forms_Unload()
    'Unload AboutForm      ' Névjegy form bezárása.
    'Unload VForm          ' Végeredmény form bezárása.
    'Unload SettingsForm   ' Beállítások form bezárása.
    'Unload WarningNewGame ' Új játék figyelmeztetés form bezárása.
    'Unload Me             ' Pálya form bezárása.
    End
End Sub

' Kiírt körök felíratának megváltoztatása.
Public Sub SetKorokSzama(KorSz As Byte)
    KorKiiras.Caption = "Kör: " & KorSz & "/" & Config.Globalis_KorokSzama
End Sub

Private Sub Timer_Korok_Timer()
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To AutokSzama
        If Korok < PalyaInfo.Autok(i).GetKorokSzama Then
            Korok = PalyaInfo.Autok(i).GetKorokSzama

            If Korok > Config.Globalis_KorokSzama Then
                VForm.Show
                Timer_Korok.Enabled = False
                Exit Sub
            End If

            SetKorokSzama Korok
        End If
    Next i
End Sub

Private Sub Timer_VersenyAdatok_Timer()
    Dim i As Byte, ciklus As Single, szam As Single, Szin As String

    Select Case VersenyAdatok.List(VersenyAdatok.ListIndex)
        Case "Autók sorrendje"
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            Dim tempkor As Byte, tempautok As Byte

            If Korok > Config.Globalis_KorokSzama Then
                tempkor = Korok - 1
            Else
                tempkor = Korok
            End If

            tempautok = 0

            Do While True
                For ciklus = 3 To 1 Step -1
                    For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To AutokSzama
                        If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                            Exit For
                        ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= AutokSzama Then
                            AddVAText i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin
                            tempautok = tempautok + 1
                        End If

                        If tempautok = AutokSzama Then
                            Exit For
                        End If
                    Next i

                    If tempautok = AutokSzama Then
                        Exit For
                    End If
                Next ciklus

                If tempautok = AutokSzama Then
                    Exit Do
                End If

                If tempkor > KezdokorErteke Then
                    tempkor = tempkor - 1
                Else
                    Exit Do
                End If
            Loop

            If tempautok = 0 Then
                AddVAText "Nincs még sorrend!"
            Else
                AddVAText ""
            End If

            AddVAText "A sorrend mindig a következõ szektornál frissül!"
        Case "Legjobb 1. szektor"
            SzektoridoKiiras 1
        Case "Legjobb 2. szektor"
            SzektoridoKiiras 2
        Case "Legjobb 3. szektor"
            SzektoridoKiiras 3
        Case "Legjobb köridõ"
            Dim lkor As Byte, aszam As Byte
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            If Korok = KezdokorErteke Then
                AddVAText "Nincs még mért köridõ!"
                Exit Sub
            End If

            szam = KezdoSzektorido
            For i = LBound(PalyaInfo.Autok) To AutokSzama
                If szam > PalyaInfo.Autok(i).GetLegjobbKorido Then
                    aszam = i
                    Szin = PalyaInfo.Autok(i).GetColor
                    szam = PalyaInfo.Autok(i).GetLegjobbKorido
                    lkor = PalyaInfo.Autok(i).GetLegjobbKoridoSzama
                End If
            Next i

            AddVAText "Legjobb kör ideje: " & szam & " másodperc"
            AddVAText "Elsõ szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " másodperc"
            AddVAText "Második szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " másodperc"
            AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " másodperc"
            AddVAText ""
            AddVAText "Az idõ a(z) " & lkor & ". körben került beállításra."
            AddVAText "A(z) idõt beállította a " & Szin & " szinû autó."
        Case "Elméleti legjobb köridõ"
            Dim T(1 To 3) As Single, TSzin(1 To 3) As String, ljekorido As Single
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            If Korok = KezdokorErteke Then
                AddVAText "Nincs még mért köridõ!"
                Exit Sub
            End If

            For i = 1 To 3
                T(i) = LegjobbSzektorido(i, TSzin(i))
                ljekorido = ljekorido + T(i)
            Next i

            AddVAText "Legjobb elméleti köriõ: " & ljekorido & " másodperc"
            AddVAText "Elsõ szektor ideje: " & T(1) & " másodperc"
            AddVAText "Autó színe: " & TSzin(1)
            AddVAText ""
            AddVAText "Második szektor ideje: " & T(2) & " másodperc"
            AddVAText "Autó színe: " & TSzin(2)
            AddVAText ""
            AddVAText "Harmadik szektor ideje: " & T(3) & " másodperc"
            AddVAText "Autó színe: " & TSzin(3)
        Case "Pálya hossza"
            CleanVAText

            If Not Started Then
                NoStartedGameVAText
                Exit Sub
            End If

            If Korok = KezdokorErteke Then
                AddVAText "Még nem sikerült megmérni a pálya hosszát!"
                Exit Sub
            End If

            AddVAText "Egy kör hossza: " & PalyaInfo.Autok(LBound(PalyaInfo.Autok)).GetEgyKorHossza & " m"
            AddVAText "Ez az érték csak egy körübelüli érték mely a mozgásváltozással van összefüggésben."
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
            CleanVAText
            AddVAText "Hiba!"
    End Select
End Sub

Private Sub EgyeniAutoKiirasok(aszam As Byte)
    Dim lkor As Byte, szam As Single, Szin As String
    CleanVAText

    If Not Started Then
        NoStartedGameVAText
        Exit Sub
    End If

    If aszam > AutokSzama Then
        AddVAText "Ez az autó nem versenyezik!"
        Exit Sub
    End If

    If Korok = KezdokorErteke Then
        AddVAText "Nincs még mért köridõ!"
        Exit Sub
    End If

    szam = KezdoSzektorido

    If szam > PalyaInfo.Autok(aszam).GetLegjobbKorido Then
        Szin = PalyaInfo.Autok(aszam).GetColor
        szam = PalyaInfo.Autok(aszam).GetLegjobbKorido
        lkor = PalyaInfo.Autok(aszam).GetLegjobbKoridoSzama
    End If

    AddVAText "Legjobb kör ideje: " & szam & " másodperc"
    AddVAText "Elsõ szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(1) & " másodperc"
    AddVAText "Második szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(2) & " másodperc"
    AddVAText "Harmadik szektor ideje: " & PalyaInfo.Autok(aszam).GetLegjobbSzektoridok(3) & " másodperc"
    AddVAText ""
    AddVAText "Az idõ a(z) " & lkor & ". körben került beállításra."
End Sub

Private Function LegjobbSzektorido(ByVal a As Integer, Szin As String) As Single
    Dim i As Byte, szam As Single
    CleanVAText
    szam = KezdoSzektorido

    For i = LBound(PalyaInfo.Autok) To AutokSzama
        If szam > PalyaInfo.Autok(i).GetLegjobbSzektoridok(a) Then
            Szin = PalyaInfo.Autok(i).GetColor
            szam = PalyaInfo.Autok(i).GetLegjobbSzektoridok(a)
        End If
    Next i

    LegjobbSzektorido = szam
End Function

Private Sub SzektoridoKiiras(a As Integer)
    Dim i As Byte, szam As Single, Szin As String
    CleanVAText

    If Not Started Then
        NoStartedGameVAText
        Exit Sub
    End If

    szam = LegjobbSzektorido(a, Szin)

    If szam = KezdoSzektorido Then ' Kezdõérték
        AddVAText "Nincs még mért szektoridõ!"
        Exit Sub
    End If

    AddVAText "Legjobb szektoridõ: " & szam & " másodperc"
    AddVAText "Az idõt beállította a " & Szin & " szinû autó."
End Sub

Private Sub NoStartedGameVAText()
    VersenyAdatokText.Text = "Még nem indult el a játék!"
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

    If TempAutoLista = AutoLista.List(AutoLista.ListIndex) Then
        Exit Sub
    End If

    Select Case AutoLista.List(AutoLista.ListIndex)
        Case "Kettõ autó"
            UjAutokLetrehozasa 2
        Case "Három autó"
            UjAutokLetrehozasa 3
        Case "Négy autó"
            UjAutokLetrehozasa 4
        Case Else
            CleanALText
            AddALText "Hiba!"
    End Select
End Sub

Public Sub UjAutokLetrehozasa(UjAutokSzama As Byte)
    TempAutoLista = AutoLista.List(AutoLista.ListIndex)
    Dispose_Game
    New_Game UjAutokSzama
    AutokKiirasa
End Sub

Private Sub AutokKiirasa()
    If Started Then
        Exit Sub
    End If

    Dim i As Byte
    CleanALText

    For i = LBound(PalyaInfo.Autok) To AutokSzama
        AddALText "[" & i & ". Autó] Színe: " & PalyaInfo.Autok(i).GetColor()
    Next i
End Sub

Private Sub CleanALText()
    AutoListaText.Text = ""
End Sub

Private Sub AddALText(Szoveg As String)
    AutoListaText.Text = AutoListaText.Text & Szoveg & vbCrLf
End Sub

Public Sub HamisPalya_Frissites()
    SetSzektorNevek
    SetSzektorVonalak

    If Not PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        PalyaInfo.StartCelVonalNev.Label.visible = Config.Globalis_StartCelVonalNeve
    End If

    If Not PalyaInfo.SzektorVonalakSzama = 0 Then
        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama - 1).Vonal.visible = Config.Globalis_StartCelVonal
    End If

    Dim vPt As POINTAPI

    With HamisPalya
       ClientToScreen .hWnd, vPt                             ' Konvertálás a 0, 0 képernyõ kordinátára.
       ScreenToClient .Container.hWnd, vPt                   ' "Container" kordináták konvertálása.
       SetViewportOrgEx .hDC, -vPt.X, -vPt.Y, vPt            ' Eltolás PictureBox DC.
       SendMessage .Container.hWnd, WM_PAINT, .hDC, ByVal 0& ' VB átfestés.
       SetViewportOrgEx .hDC, vPt.X, vPt.Y, vPt              ' Reset PictureBox DC.
       If .AutoRedraw = True Then .Refresh
    End With
End Sub
