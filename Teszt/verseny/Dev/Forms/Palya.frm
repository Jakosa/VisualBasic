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
      TabIndex        =   9
      Top             =   600
      Width           =   9855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Autók"
      Height          =   3495
      Left            =   10440
      TabIndex        =   4
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
         TabIndex        =   7
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox AutoLista 
         Height          =   315
         ItemData        =   "Palya.frx":0000
         Left            =   240
         List            =   "Palya.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Versenyadatok"
      Height          =   3615
      Left            =   10440
      TabIndex        =   2
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
         TabIndex        =   6
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox VersenyAdatok 
         Height          =   315
         ItemData        =   "Palya.frx":0034
         Left            =   240
         List            =   "Palya.frx":004D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
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
      TabIndex        =   10
      Top             =   120
      Width           =   1095
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
      Index           =   13
      X1              =   960
      X2              =   720
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Start / Cél / Szektor 3"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   4440
      Width           =   1560
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Szektor 2"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Szektor 1"
      Height          =   195
      Left            =   4680
      TabIndex        =   0
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
      Index           =   1
      X1              =   720
      X2              =   720
      Y1              =   4920
      Y2              =   5520
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
      Begin VB.Menu SaveResult 
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
         Checked         =   -1  'True
         Shortcut        =   ^N
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
Dim WithEvents Timer_VersenyAdatok As VB.Timer  ' VersenyAdatok elnevezésõ listát frissíti.
Attribute Timer_VersenyAdatok.VB_VarHelpID = -1
Dim WithEvents Timer_AutoLista As VB.Timer      ' AutoLista elnevezésõ listát frissíti.
Attribute Timer_AutoLista.VB_VarHelpID = -1
Dim WithEvents Timer_Korok As VB.Timer          ' Frissíti a körök számát. (Ha új kör van megváltoztatja a számlálót is.)
Attribute Timer_Korok.VB_VarHelpID = -1
Dim AutokSzama As Byte                          ' Versenypályán lévõ autók számát tárolja.
Dim Korok As Byte                               ' Tárolja éppen hányadik körnél tartunk.
Dim Started As Boolean                          ' Jelzi hogy elindult-e már a játék vagy sem.
Public TempAutoLista As String                  ' Hány autó van kiválasztva. (terheléscsökkentés)
Dim Felfuggesztes As Boolean                    ' Ha használva van a Stop gomb akkor lesz "true" az értéke.
Const KezdokorErteke = 1                        ' Tárolja hogy hánytól induljon az elsõ kör.
Const BorderWidth = 2                           ' Autók vonalának szélessége.
Const PalyaHosszanakLepteke = 10                ' 10 m-t jelent. Ez azt jelenti hogy egy elmozdulással az autó 10 métert tesz meg.
Const ex = 0.6
Const ey = -1

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
    Timer_AutoLista.Interval = 500     ' Érték beállítása. 500 millisec

    ' Nyomvonal megjelenésének beállítása
    Nyomvonal.Checked = Config.Globalis_Nyomvonal

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
    T(3) = "sárga"
    T(4) = "zöld"

    Dim tavolodas As Integer
    For i = 1 To ASzama
        Autok(i).Load i ' Betöltjük újként a vonalat
        Autok(i).SetEX ex
        Autok(i).SetEY ey
        Autok(i).SetX0 1100 - tavolodas
        Autok(i).SetY0 5000
        Autok(i).SetColor T(i) ' Ha kell színezés csak akkor.
        Autok(i).SetBorderWidth BorderWidth
        Autok(i).Show
        tavolodas = tavolodas + 30
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
        Autok(i).Dispose
        Set Autok(i) = Nothing
    Next i

    AutokSzama = 0
End Sub

Private Sub GlobalSettings_Click()
    If Started Then
        MsgBox "A játék már fut! Indíts új játékot ha szeretnél a beállíátsokon változtatni.", , "Beállítások: Hiba!"
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
    For i = LBound(Autok) To AutokSzama
        If Not Autok(i).GetGameEnd Then
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
    ReDim SorrendTomb(KezdokorErteke To Config.Globalis_KorokSzama) As Sorrend
End Sub

Private Sub Nyomvonal_Click()
    If Nyomvonal.Checked Then
        Nyomvonal.Checked = False
    Else
        Nyomvonal.Checked = True
    End If

    Config.Globalis_Nyomvonal = Nyomvonal.Checked
    Config.SetConfig

    ' Itt kell majd az autok nyomvonalát buherálni......
    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        Autok(i).SetNyomvonal Nyomvonal.Checked
    Next i
End Sub

Private Sub Start_Click()
    If AutokSzama = 0 Then
        MsgBox "Még nincsenek kiválasztva autók!"
        Exit Sub
    End If

    If Not Started Then
        Started = True
    End If

    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        If Autok(i).GetGameEnd Then
            MsgBox "A játék végetért! Nem indíthatod már el Start-tal! Indíts új játékot ha újat kezdenél."
            Exit Sub
        End If
    Next i

    If Felfuggesztes Then
        MsgBox "A játék már fut!", , "Start: Hiba!"
        Exit Sub
    End If

    If Not Felfuggesztes Then
        Felfuggesztes = True
    End If

    For i = LBound(Autok) To AutokSzama
        Autok(i).Start
    Next i
End Sub

Private Sub Stop_Click()
    If Not Felfuggesztes Then
        MsgBox "A játék nem fut!", , "Stop: Hiba!"
        Exit Sub
    End If

    If Felfuggesztes Then
        Felfuggesztes = False
    End If

    Dim i As Byte
    For i = LBound(Autok) To AutokSzama
        Autok(i).Stop_Kocsi
    Next i
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
    For i = LBound(Autok) To AutokSzama
        If Korok < Autok(i).GetKorokSzama Then
            Korok = Autok(i).GetKorokSzama

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
                    For i = LBound(SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To AutokSzama
                        If SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                            Exit For
                        ElseIf SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= AutokSzama Then
                            AddVAText i & ". Autó: " & SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin
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

            szam = 1000
            For i = LBound(Autok) To AutokSzama
                If szam > Autok(i).GetLegjobbKorido Then
                    aszam = i
                    Szin = Autok(i).GetColor
                    szam = Autok(i).GetLegjobbKorido
                    lkor = Autok(i).GetLegjobbKoridoSzama
                End If
            Next i

            AddVAText "Legjobb kör ideje: " & szam & " másodperc"
            AddVAText "Elsõ szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(1) & " másodperc"
            AddVAText "Második szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(2) & " másodperc"
            AddVAText "Harmadik szektor ideje: " & Autok(aszam).GetLegjobbSzektoridok(3) & " másodperc"
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

            AddVAText "Egy kör hossza: " & Autok(LBound(Autok)).GetEgyKorHossza & " m"
            AddVAText "Ez az érték csak egy körübelüli érték mely a mozgásváltozással van összefüggésben."
            AddVAText PalyaHosszanakLepteke & " egység (egy lépés) felel meg " & PalyaHosszanakLepteke & " méternek."
        Case Else
            CleanVAText
            AddVAText "Hiba!"
    End Select
End Sub

Private Function LegjobbSzektorido(ByVal a As Integer, Szin As String) As Single
    Dim i As Byte, szam As Single
    CleanVAText
    szam = 1000

    For i = LBound(Autok) To AutokSzama
        If szam > Autok(i).GetLegjobbSzektoridok(a) Then
            Szin = Autok(i).GetColor
            szam = Autok(i).GetLegjobbSzektoridok(a)
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

    If szam = 10000000 Then ' Kezdõérték
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
            TempAutoLista = AutoLista.List(AutoLista.ListIndex)
            Dispose_Game
            New_Game 2
            AutokKiirasa
        Case "Három autó"
            TempAutoLista = AutoLista.List(AutoLista.ListIndex)
            Dispose_Game
            New_Game 3
            AutokKiirasa
        Case "Négy autó"
            TempAutoLista = AutoLista.List(AutoLista.ListIndex)
            Dispose_Game
            New_Game 4
            AutokKiirasa
        Case Else
            CleanALText
            AddALText "Hiba!"
    End Select
End Sub

Private Sub AutokKiirasa()
    If Started Then
        Exit Sub
    End If

    Dim i As Byte
    CleanALText

    For i = LBound(Autok) To AutokSzama
        AddALText "[" & i & ". Autó] Színe: " & Autok(i).GetColor()
    Next i
End Sub

Private Sub CleanALText()
    AutoListaText.Text = ""
End Sub

Private Sub AddALText(Szoveg As String)
    AutoListaText.Text = AutoListaText.Text & Szoveg & vbCrLf
End Sub

Private Sub HamisPalya_Frissites()
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

Private Function Distance(ByVal PointX As Single, ByVal PointY As Single, ByVal LineX1 As Single, ByVal LineX2 As Single, ByVal LineY1 As Single, ByVal LineY2 As Single) As Single
    Dim AA As Single, BB As Single, CC As Single, DD As Single
    Dim dot As Single, len_sq As Single, param As Single
    Dim xx As Single, yy As Single
    AA = PointX - LineX1
    BB = PointY - LineY1
    CC = LineX2 - LineX1
    DD = LineY2 - LineY1

    dot = AA * CC + BB * DD
    len_sq = CC * CC + DD * DD
    param = dot / len_sq

    If param < 0 Then
        xx = LineX1
        yy = LineY1
    ElseIf param > 1 Then
        xx = LineX2
        yy = LineY2
    Else
        xx = LineX1 + param * CC
        yy = LineY1 + param * DD
    End If

    Distance = Sqr(((PointX - xx) * (PointX - xx)) + ((PointY - yy) * (PointY - yy)))
End Function
