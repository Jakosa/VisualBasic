VERSION 5.00
Begin VB.Form SettingsForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beállítások"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame NegyedikFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Negyedik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   32
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckNegyedikNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox CheckNegyedikTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tökéletes körözés be illetve kikapcsolásának lehetõsége."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Frame HarmadikFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Harmadik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   29
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckHarmadikTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tökéletes körözés be illetve kikapcsolásának lehetõsége."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   5175
      End
      Begin VB.CheckBox CheckHarmadikNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame MasodikFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Második"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   26
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckMasodikNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox CheckMasodikTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tökéletes körözés be illetve kikapcsolásának lehetõsége."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Frame ElsoFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Elsõ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   23
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckElsoTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tökéletes körözés be illetve kikapcsolásának lehetõsége."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   5175
      End
      Begin VB.CheckBox CheckElsoNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame PalyaFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pálya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   18
      Top             =   120
      Width           =   9375
      Begin VB.ComboBox PalyaComboBox 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pálya kiválasztása:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pályához tartozó körök száma:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   2730
      End
      Begin VB.Label PTKorokSzama 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Altalanos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Általános"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tökéletes körözés be illetve kikapcsolásának lehetõsége."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   4695
      End
      Begin VB.CheckBox CheckStartCelVonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start és célvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CheckBox CheckSzektorVonalak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Szektor vonalak be illetve kikapcsolásának lehetõsége."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   4935
      End
      Begin VB.CheckBox CheckStartCelVonalNeve 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A start és célvonal neve be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   5055
      End
      Begin VB.CheckBox CheckSzektorNevek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Szektor nevek be illetve kikapcsolásának lehetõsége."
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   4335
      End
      Begin VB.ComboBox KorokComboBox 
         Height          =   315
         ItemData        =   "SettingsForm.frx":0000
         Left            =   1920
         List            =   "SettingsForm.frx":003D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox CheckNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsolásának lehetõsége."
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Körök száma:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdMegse 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdAlkalmaz 
      Caption         =   "Alkalmaz"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdAlapertelmezes 
      Caption         =   "Alapértelmezés"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Kategoriak 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategóriák"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.ListBox AutokLista 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         ItemData        =   "SettingsForm.frx":0095
         Left            =   480
         List            =   "SettingsForm.frx":00A5
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox GlobalisLista 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         ItemData        =   "SettingsForm.frx":00CC
         Left            =   480
         List            =   "SettingsForm.frx":00D6
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Autók"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Globális"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   750
      End
   End
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fejléc
' Készítette: Belsõ Vazul István
' Fejléc vége

Option Explicit

' Ideiglenes változó.
Private TempGlobalis_Nyomvonal As Boolean
' Ideiglenes változó.
Private TempGlobalis_SzektorNevek As Boolean
' Ideiglenes változó.
Private TempGlobalis_StartCelVonalNeve As Boolean
' Ideiglenes változó.
Private TempGlobalis_KorokSzama As Byte
' Ideiglenes változó.
Private TempGlobalis_PalyaNeve As String
' Ideiglenes változó.
Private TempGlobalis_SzektorVonalak As Boolean
' Ideiglenes változó.
Private TempGlobalis_StartCelVonal As Boolean
' Ideiglenes változó.
Private TempGlobalis_TokeletesKorozes As Boolean
' Ideiglenes változó.
Private TempAutok_Elso_Nyomvonal As Boolean
' Ideiglenes változó.
Private TempAutok_Elso_TokeletesKorozes As Boolean
' Ideiglenes változó.
Private TempAutok_Masodik_Nyomvonal As Boolean
' Ideiglenes változó.
Private TempAutok_Masodik_TokeletesKorozes As Boolean
' Ideiglenes változó.
Private TempAutok_Harmadik_Nyomvonal As Boolean
' Ideiglenes változó.
Private TempAutok_Harmadik_TokeletesKorozes As Boolean
' Ideiglenes változó.
Private TempAutok_Negyedik_Nyomvonal As Boolean
' Ideiglenes változó.
Private TempAutok_Negyedik_TokeletesKorozes As Boolean

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Form_Load()
    ' Tárolja a ListBox stílusát.
    Dim lStyle As Long

    ' ListBox keretének eltávolítása.
    lStyle = GetWindowLong(GlobalisLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(GlobalisLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(GlobalisLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    ' Kezdõelem beállítása.
    GlobalisLista.ListIndex = 0

    ' ListBox keretének eltávolítása.
    lStyle = GetWindowLong(AutokLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(AutokLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(AutokLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    ' Beállítások inicializálása.
    Init
End Sub

' CmdOk gomb eseménye kattintás hatására.
Private Sub CmdOk_Click()
    ' CmdAlkalmaz_Click eljárás meghívása.
    CmdAlkalmaz_Click
    ' Form bezárása.
    Unload Me
End Sub

' CmdMegse gomb eseménye kattintás hatására.
Private Sub CmdMegse_Click()
    ' Form bezárása.
    Unload Me
End Sub

' CmdAlkalmaz gomb eseménye kattintás hatására.
Private Sub CmdAlkalmaz_Click()
    ' Ideiglenes változó ami a pálya nevét tárolja.
    Dim tname As String
    ' Pálya nevének átvétele.
    tname = Config.Globalis_PalyaNeve

    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_Nyomvonal = TempGlobalis_Nyomvonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_SzektorNevek = TempGlobalis_SzektorNevek
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_StartCelVonalNeve = TempGlobalis_StartCelVonalNeve
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_KorokSzama = TempGlobalis_KorokSzama
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_PalyaNeve = TempGlobalis_PalyaNeve
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_SzektorVonalak = TempGlobalis_SzektorVonalak
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_StartCelVonal = TempGlobalis_StartCelVonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Globalis_TokeletesKorozes = TempGlobalis_TokeletesKorozes
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Elso_Nyomvonal = TempAutok_Elso_Nyomvonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Elso_TokeletesKorozes = TempAutok_Elso_TokeletesKorozes
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Masodik_Nyomvonal = TempAutok_Masodik_Nyomvonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Masodik_TokeletesKorozes = TempAutok_Masodik_TokeletesKorozes
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Harmadik_Nyomvonal = TempAutok_Harmadik_Nyomvonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Harmadik_TokeletesKorozes = TempAutok_Harmadik_TokeletesKorozes
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Negyedik_Nyomvonal = TempAutok_Negyedik_Nyomvonal
    ' Ideiglenes beállítás betöltése a konfig beállításba.
    Config.Autok_Negyedik_TokeletesKorozes = TempAutok_Negyedik_TokeletesKorozes
    ' Konfig fájl beállítása.
    Config.SetConfig

    ' Megváltoztatjuk a körök számának kiírást.
    Palya.SetKorokSzama Palya.GetKezdokorErteke
    ' Új autók létrehozása.
    Palya.UjAutokLetrehozasa PalyaInfo.AutokSzama

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_Nyomvonal Then
        ' Bekapcsolja a pipát.
        Palya.Nyomvonal.Checked = 1
    Else
        ' Kikapcsolja a pipát.
        Palya.Nyomvonal.Checked = 0
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_TokeletesKorozes Then
        ' Bekapcsolja a pipát.
        Palya.Tokeletes_Korozes.Checked = 1
    Else
        ' Kikapcsolja a pipát.
        Palya.Tokeletes_Korozes.Checked = 0
    End If

    ' Pálya betöltése.
    Map.LoadMap Config.Globalis_PalyaNeve

    ' Akkor fut le ha valamely adatok hibásak.
    If Not Vizsgalat Then
        ' Form bezárása.
        Unload Me
    End If

    ' Akkor fut le ha az ideiglenes pálya nem egyezne meg a konfigban tárolt pálya nevével.
    If Not tname = Config.Globalis_PalyaNeve Then
        ' Körök számának beállítása.
        Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
        ' Konfig fájl beállítása.
        Config.SetConfig

        ' Kezdõelem beállítása.
        KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
        ' Megváltoztatjuk a körök számának kiírást.
        Palya.SetKorokSzama Palya.GetKezdokorErteke
        ' Új autók létrehozása.
        Palya.UjAutokLetrehozasa PalyaInfo.AutokSzama
    End If

    ' Frissítési az autók nyomvonalának megjelenítését.
    Palya.SetAutokNyomvonal
    ' Frissítési az autók tökéletes körözésének állapotát.
    Palya.SetAutokTokeletesKorozes
    ' Pályához tartozó kör számának kiírása.
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    ' Frissíti a virtuálisan létrehozott pályát.
    Palya.VirtualisPalya_Frissites
End Sub

Private Sub CmdAlapertelmezes_Click()
    ' Konfig fájl törlése.
    Config.DeleteConfig
    ' Konfig fájl betöltése.
    Config.LoadConfig

    ' Alapértelmezett pálya törlése.
    Map.DeleteDefaultMap
    ' Pálya betöltése.
    Map.LoadMap Config.Globalis_PalyaNeve
    ' Körök számának beállítása.
    Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
    ' Konfig fájl beállítása.
    Config.SetConfig

    ' Beállítások inicializálása.
    Init
End Sub

' CheckNyomvonal gomb eseménye kattintás hatására.
Private Sub CheckNyomvonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckNyomvonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_Nyomvonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckNyomvonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_Nyomvonal = False
    End If
End Sub

' CheckSzektorNevek gomb eseménye kattintás hatására.
Private Sub CheckSzektorNevek_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckSzektorNevek.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_SzektorNevek = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckSzektorNevek.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_SzektorNevek = False
    End If
End Sub

' CheckStartCelVonalNeve gomb eseménye kattintás hatására.
Private Sub CheckStartCelVonalNeve_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckStartCelVonalNeve.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_StartCelVonalNeve = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckStartCelVonalNeve.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_StartCelVonalNeve = False
    End If
End Sub

' CheckSzektorVonalak gomb eseménye kattintás hatására.
Private Sub CheckSzektorVonalak_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckSzektorVonalak.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_SzektorVonalak = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckSzektorVonalak.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_SzektorVonalak = False
    End If
End Sub

' CheckStartCelVonal gomb eseménye kattintás hatására.
Private Sub CheckStartCelVonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckStartCelVonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_StartCelVonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckStartCelVonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_StartCelVonal = False
    End If
End Sub

' CheckTokeletesKorozes gomb eseménye kattintás hatására.
Private Sub CheckTokeletesKorozes_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckTokeletesKorozes.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempGlobalis_TokeletesKorozes = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckTokeletesKorozes.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempGlobalis_TokeletesKorozes = False
    End If
End Sub

' CheckElsoNyomvonal gomb eseménye kattintás hatására.
Private Sub CheckElsoNyomvonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckElsoNyomvonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Elso_Nyomvonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckElsoNyomvonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Elso_Nyomvonal = False
    End If
End Sub

' CheckElsoTokeletesKorozes gomb eseménye kattintás hatására.
Private Sub CheckElsoTokeletesKorozes_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckElsoTokeletesKorozes.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Elso_TokeletesKorozes = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckElsoTokeletesKorozes.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Elso_TokeletesKorozes = False
    End If
End Sub

' CheckMasodikNyomvonal gomb eseménye kattintás hatására.
Private Sub CheckMasodikNyomvonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckMasodikNyomvonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Masodik_Nyomvonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckMasodikNyomvonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Masodik_Nyomvonal = False
    End If
End Sub

' CheckMasodikTokeletesKorozes gomb eseménye kattintás hatására.
Private Sub CheckMasodikTokeletesKorozes_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckMasodikTokeletesKorozes.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Masodik_TokeletesKorozes = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckMasodikTokeletesKorozes.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Masodik_TokeletesKorozes = False
    End If
End Sub

' CheckHarmadikNyomvonal gomb eseménye kattintás hatására.
Private Sub CheckHarmadikNyomvonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckHarmadikNyomvonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Harmadik_Nyomvonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckHarmadikNyomvonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Harmadik_Nyomvonal = False
    End If
End Sub

' CheckHarmadikTokeletesKorozes gomb eseménye kattintás hatására.
Private Sub CheckHarmadikTokeletesKorozes_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckHarmadikTokeletesKorozes.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Harmadik_TokeletesKorozes = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckHarmadikTokeletesKorozes.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Harmadik_TokeletesKorozes = False
    End If
End Sub

' CheckNegyedikNyomvonal gomb eseménye kattintás hatására.
Private Sub CheckNegyedikNyomvonal_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckNegyedikNyomvonal.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Negyedik_Nyomvonal = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckNegyedikNyomvonal.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Negyedik_Nyomvonal = False
    End If
End Sub

' CheckNegyedikTokeletesKorozes gomb eseménye kattintás hatására.
Private Sub CheckNegyedikTokeletesKorozes_Click()
    ' Ha egyenlõ eggyel akkor fut le.
    If CheckNegyedikTokeletesKorozes.value = 1 Then
        ' Igazra állítja az ideiglenes változót.
        TempAutok_Negyedik_TokeletesKorozes = True
    ' Ha egyenlõ nullával akkor fut le.
    ElseIf CheckNegyedikTokeletesKorozes.value = 0 Then
        ' Hamisra állítja az ideiglenes változót.
        TempAutok_Negyedik_TokeletesKorozes = False
    End If
End Sub

' KorokComboBox lista eseménye kattintás hatására.
Private Sub KorokComboBox_Click()
    ' Ideiglenes körök számának megváltoztatása az index alapján.
    TempGlobalis_KorokSzama = CByte(Trim(KorokComboBox.List(KorokComboBox.ListIndex)))
End Sub

' PalyaComboBox lista eseménye kattintás hatására.
Private Sub PalyaComboBox_Click()
    ' Ideiglenes pálya nevének megváltoztatása az index alapján.
    TempGlobalis_PalyaNeve = PalyaComboBox.List(PalyaComboBox.ListIndex)
End Sub

' Globális listán kiválasztjuk a megjelenitendõ "frame"-t.
Private Sub GlobalisLista_Click()
    ' "Frame"-k láthatatlanná tétele.
    SetAllVisible False

    ' Kiválasztott elem alapján szelektálja melyik "frame" jelenjen meg.
    Select Case GlobalisLista.List(GlobalisLista.ListIndex)
        Case "Általános"
            ' Autók lista kijelölésének eltüntetése.
            AutokLista.ListIndex = -1
            ' "Frame" megjelenítése.
            Altalanos.Visible = True
        Case "Pálya"
            ' Autók lista kijelölésének eltüntetése.
            AutokLista.ListIndex = -1
            ' "Frame" megjelenítése.
            PalyaFrame.Visible = True
    End Select
End Sub

' Autók listáján kiválasztjuk a megjelenitendõ "frame"-t.
Private Sub AutokLista_Click()
    ' "Frame"-k láthatatlanná tétele.
    SetAllVisible False

    ' Kiválasztott elem alapján szelektálja melyik "frame" jelenjen meg.
    Select Case AutokLista.List(AutokLista.ListIndex)
        Case "Elsõ"
            ' Globális lista kijelölésének eltüntetése.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelenítése.
            ElsoFrame.Visible = True
        Case "Második"
            ' Globális lista kijelölésének eltüntetése.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelenítése.
            MasodikFrame.Visible = True
        Case "Harmadik"
            ' Globális lista kijelölésének eltüntetése.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelenítése.
            HarmadikFrame.Visible = True
        Case "Negyedik"
            ' Globális lista kijelölésének eltüntetése.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelenítése.
            NegyedikFrame.Visible = True
    End Select
End Sub

' Alapértelmezett folyamatok, beállítások betöltése.
Private Sub Init()
    ' Pályaadatok beállítása.
    SetPalyaComboBox

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_Nyomvonal Then
        ' Bekapcsolja a pipát.
        CheckNyomvonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    Else
        ' Kikapcsolja a pipát.
        CheckNyomvonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_SzektorNevek Then
        ' Bekapcsolja a pipát.
        CheckSzektorNevek.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    Else
        ' Kikapcsolja a pipát.
        CheckSzektorNevek.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_StartCelVonalNeve Then
        ' Bekapcsolja a pipát.
        CheckStartCelVonalNeve.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    Else
        ' Kikapcsolja a pipát.
        CheckStartCelVonalNeve.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_SzektorVonalak Then
        ' Bekapcsolja a pipát.
        CheckSzektorVonalak.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    Else
        ' Kikapcsolja a pipát.
        CheckSzektorVonalak.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_StartCelVonal Then
        ' Bekapcsolja a pipát.
        CheckStartCelVonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    Else
        ' Kikapcsolja a pipát.
        CheckStartCelVonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Globalis_TokeletesKorozes Then
        ' Bekapcsolja a pipát.
        CheckTokeletesKorozes.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    Else
        ' Kikapcsolja a pipát.
        CheckTokeletesKorozes.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Elso_Nyomvonal Then
        ' Bekapcsolja a pipát.
        CheckElsoNyomvonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    Else
        ' Kikapcsolja a pipát.
        CheckElsoNyomvonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Elso_TokeletesKorozes Then
        ' Bekapcsolja a pipát.
        CheckElsoTokeletesKorozes.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    Else
        ' Kikapcsolja a pipát.
        CheckElsoTokeletesKorozes.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Masodik_Nyomvonal Then
        ' Bekapcsolja a pipát.
        CheckMasodikNyomvonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    Else
        ' Kikapcsolja a pipát.
        CheckMasodikNyomvonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Masodik_TokeletesKorozes Then
        ' Bekapcsolja a pipát.
        CheckMasodikTokeletesKorozes.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    Else
        ' Kikapcsolja a pipát.
        CheckMasodikTokeletesKorozes.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Harmadik_Nyomvonal Then
        ' Bekapcsolja a pipát.
        CheckHarmadikNyomvonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    Else
        ' Kikapcsolja a pipát.
        CheckHarmadikNyomvonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Harmadik_TokeletesKorozes Then ' Kikapcsolva!!!
        ' Bekapcsolja a pipát.
        CheckHarmadikTokeletesKorozes.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    Else
        ' Kikapcsolja a pipát.
        CheckHarmadikTokeletesKorozes.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Negyedik_Nyomvonal Then
        ' Bekapcsolja a pipát.
        CheckNegyedikNyomvonal.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    Else
        ' Kikapcsolja a pipát.
        CheckNegyedikNyomvonal.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    End If

    ' Ha igaz az érték akkor fut le.
    If Config.Autok_Negyedik_TokeletesKorozes Then
        ' Bekapcsolja a pipát.
        CheckNegyedikTokeletesKorozes.value = 1
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    Else
        ' Kikapcsolja a pipát.
        CheckNegyedikTokeletesKorozes.value = 0
        ' Ideiglenes beállítás módosítása a konfigban tároltra.
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    End If

    ' Frissítési az autók nyomvonalának megjelenítését.
    Palya.SetAutokNyomvonal
    ' Frissítési az autók tökéletes körözésének állapotát.
    Palya.SetAutokTokeletesKorozes
    ' Ideiglenes beállítás módosítása a konfigban tároltra.
    TempGlobalis_PalyaNeve = Config.Globalis_PalyaNeve
    ' Kezdõelem beállítása.
    KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
    ' Ideiglenes beállítás módosítása a konfigban tároltra.
    TempGlobalis_KorokSzama = Config.Globalis_KorokSzama
    ' Megváltoztatjuk a körök számának kiírást.
    Palya.SetKorokSzama Palya.GetKezdokorErteke
    ' Pályához tartozó kör számának kiírása.
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    ' Frissíti a virtuálisan létrehozott pályát.
    Palya.VirtualisPalya_Frissites
End Sub

' Összes "frame" láthatatlanná vagy láthatóvá tétele.
' A "Visible" tárolja a "frame"-k láthatóságát.
Private Sub SetAllVisible(ByVal Visible As Boolean)
    ' Láthatóság beállítása.
    Altalanos.Visible = Visible
    ' Láthatóság beállítása.
    PalyaFrame.Visible = Visible
    ' Láthatóság beállítása.
    ElsoFrame.Visible = Visible
    ' Láthatóság beállítása.
    MasodikFrame.Visible = Visible
    ' Láthatóság beállítása.
    HarmadikFrame.Visible = Visible
    ' Láthatóság beállítása.
    NegyedikFrame.Visible = Visible
End Sub

Private Sub SetPalyaComboBox()
    ' Könyvtár adatok tárolása.
    Dim a As String
    ' Tárolja pálya nevének indexét.
    Dim index As Integer
    ' Mappa alkönyvtárainak és fájlainak betöltése.
    a = Dir$(MapDir & "\*.*", vbDirectory)
    ' Pályak számát nulláza.
    PalyaInfo.PalyaNevekSzama = 0
    ' Törlis a pályák listáját.
    PalyaComboBox.Clear
    ' Törlis a pálya neveket a tömbbõl.
    ReDim PalyaInfo.PalyaNevek(0 To 10) As String

    ' Addig fut amíg nem egyenlõ az "a" a semmivel.
    Do While a <> ""
        ' Megnézi hogy fájl-e.
        If IsFile(MapDir & "\" & a) Then
            ' Ha nagyobb a pályák száma mint amit a tömb tárolni tud akkor fut le.
            If PalyaInfo.PalyaNevekSzama >= UBound(PalyaInfo.PalyaNevek) Then
                ' Tömb megnövelése.
                ReDim Preserve PalyaInfo.PalyaNevek(0 To UBound(PalyaInfo.PalyaNevek) + 10) As String
            End If

            ' Megnézi hogy a konfigban tárolt pálya neve megegyezik-e a betöltött névvel.
            If a = Config.Globalis_PalyaNeve Then
                ' Elmenti a tömbbõl az adott indexet.
                index = PalyaInfo.PalyaNevekSzama
            End If

            ' Kiírja a pálya nevét.
            PalyaComboBox.AddItem a
            ' Elmenti a pálya nevét a tömbbe.
            PalyaInfo.PalyaNevek(PalyaInfo.PalyaNevekSzama) = a
            ' Megnöveli a pálya nevek számát.
            PalyaInfo.PalyaNevekSzama = PalyaInfo.PalyaNevekSzama + 1
        End If

        ' Könyvtár ugrás.
        a = Dir$
    Loop

    ' Kezdõelem beállítása.
    PalyaComboBox.ListIndex = index
End Sub
