VERSION 5.00
Begin VB.Form SettingsForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Be�ll�t�sok"
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
         Caption         =   "Nyomvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox CheckNegyedikTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T�k�letes k�r�z�s be illetve kikapcsol�s�nak lehet�s�ge."
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
         Caption         =   "T�k�letes k�r�z�s be illetve kikapcsol�s�nak lehet�s�ge."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   5175
      End
      Begin VB.CheckBox CheckHarmadikNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame MasodikFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "M�sodik"
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
         Caption         =   "Nyomvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox CheckMasodikTokeletesKorozes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T�k�letes k�r�z�s be illetve kikapcsol�s�nak lehet�s�ge."
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
      Caption         =   "Els�"
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
         Caption         =   "T�k�letes k�r�z�s be illetve kikapcsol�s�nak lehet�s�ge."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   5175
      End
      Begin VB.CheckBox CheckElsoNyomvonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nyomvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame PalyaFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P�lya"
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
         Caption         =   "P�lya kiv�laszt�sa:"
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
         Caption         =   "P�ly�hoz tartoz� k�r�k sz�ma:"
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
      Caption         =   "�ltal�nos"
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
         Caption         =   "T�k�letes k�r�z�s be illetve kikapcsol�s�nak lehet�s�ge."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   4695
      End
      Begin VB.CheckBox CheckStartCelVonal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start �s c�lvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CheckBox CheckSzektorVonalak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Szektor vonalak be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   4935
      End
      Begin VB.CheckBox CheckStartCelVonalNeve 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A start �s c�lvonal neve be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   5055
      End
      Begin VB.CheckBox CheckSzektorNevek 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Szektor nevek be illetve kikapcsol�s�nak lehet�s�ge."
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
         Caption         =   "Nyomvonal be illetve kikapcsol�s�nak lehet�s�ge."
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "K�r�k sz�ma:"
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
      Caption         =   "M�gse"
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
      Caption         =   "Alap�rtelmez�s"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Kategoriak 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kateg�ri�k"
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
         Caption         =   "Aut�k"
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
         Caption         =   "Glob�lis"
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
' Fejl�c
' K�sz�tette: Bels� Vazul Istv�n
' Fejl�c v�ge

Option Explicit

' Ideiglenes v�ltoz�.
Private TempGlobalis_Nyomvonal As Boolean
' Ideiglenes v�ltoz�.
Private TempGlobalis_SzektorNevek As Boolean
' Ideiglenes v�ltoz�.
Private TempGlobalis_StartCelVonalNeve As Boolean
' Ideiglenes v�ltoz�.
Private TempGlobalis_KorokSzama As Byte
' Ideiglenes v�ltoz�.
Private TempGlobalis_PalyaNeve As String
' Ideiglenes v�ltoz�.
Private TempGlobalis_SzektorVonalak As Boolean
' Ideiglenes v�ltoz�.
Private TempGlobalis_StartCelVonal As Boolean
' Ideiglenes v�ltoz�.
Private TempGlobalis_TokeletesKorozes As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Elso_Nyomvonal As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Elso_TokeletesKorozes As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Masodik_Nyomvonal As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Masodik_TokeletesKorozes As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Harmadik_Nyomvonal As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Harmadik_TokeletesKorozes As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Negyedik_Nyomvonal As Boolean
' Ideiglenes v�ltoz�.
Private TempAutok_Negyedik_TokeletesKorozes As Boolean

' Be�ll�tjuk a form l�trehoz�sakor az alap folyamatokat.
Private Sub Form_Load()
    ' T�rolja a ListBox st�lus�t.
    Dim lStyle As Long

    ' ListBox keret�nek elt�vol�t�sa.
    lStyle = GetWindowLong(GlobalisLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(GlobalisLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(GlobalisLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    ' Kezd�elem be�ll�t�sa.
    GlobalisLista.ListIndex = 0

    ' ListBox keret�nek elt�vol�t�sa.
    lStyle = GetWindowLong(AutokLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(AutokLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(AutokLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    ' Be�ll�t�sok inicializ�l�sa.
    Init
End Sub

' CmdOk gomb esem�nye kattint�s hat�s�ra.
Private Sub CmdOk_Click()
    ' CmdAlkalmaz_Click elj�r�s megh�v�sa.
    CmdAlkalmaz_Click
    ' Form bez�r�sa.
    Unload Me
End Sub

' CmdMegse gomb esem�nye kattint�s hat�s�ra.
Private Sub CmdMegse_Click()
    ' Form bez�r�sa.
    Unload Me
End Sub

' CmdAlkalmaz gomb esem�nye kattint�s hat�s�ra.
Private Sub CmdAlkalmaz_Click()
    ' Ideiglenes v�ltoz� ami a p�lya nev�t t�rolja.
    Dim tname As String
    ' P�lya nev�nek �tv�tele.
    tname = Config.Globalis_PalyaNeve

    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_Nyomvonal = TempGlobalis_Nyomvonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_SzektorNevek = TempGlobalis_SzektorNevek
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_StartCelVonalNeve = TempGlobalis_StartCelVonalNeve
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_KorokSzama = TempGlobalis_KorokSzama
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_PalyaNeve = TempGlobalis_PalyaNeve
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_SzektorVonalak = TempGlobalis_SzektorVonalak
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_StartCelVonal = TempGlobalis_StartCelVonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Globalis_TokeletesKorozes = TempGlobalis_TokeletesKorozes
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Elso_Nyomvonal = TempAutok_Elso_Nyomvonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Elso_TokeletesKorozes = TempAutok_Elso_TokeletesKorozes
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Masodik_Nyomvonal = TempAutok_Masodik_Nyomvonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Masodik_TokeletesKorozes = TempAutok_Masodik_TokeletesKorozes
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Harmadik_Nyomvonal = TempAutok_Harmadik_Nyomvonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Harmadik_TokeletesKorozes = TempAutok_Harmadik_TokeletesKorozes
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Negyedik_Nyomvonal = TempAutok_Negyedik_Nyomvonal
    ' Ideiglenes be�ll�t�s bet�lt�se a konfig be�ll�t�sba.
    Config.Autok_Negyedik_TokeletesKorozes = TempAutok_Negyedik_TokeletesKorozes
    ' Konfig f�jl be�ll�t�sa.
    Config.SetConfig

    ' Megv�ltoztatjuk a k�r�k sz�m�nak ki�r�st.
    Palya.SetKorokSzama Palya.GetKezdokorErteke
    ' �j aut�k l�trehoz�sa.
    Palya.UjAutokLetrehozasa PalyaInfo.AutokSzama

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        Palya.Nyomvonal.Checked = 1
    Else
        ' Kikapcsolja a pip�t.
        Palya.Nyomvonal.Checked = 0
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_TokeletesKorozes Then
        ' Bekapcsolja a pip�t.
        Palya.Tokeletes_Korozes.Checked = 1
    Else
        ' Kikapcsolja a pip�t.
        Palya.Tokeletes_Korozes.Checked = 0
    End If

    ' P�lya bet�lt�se.
    Map.LoadMap Config.Globalis_PalyaNeve

    ' Akkor fut le ha valamely adatok hib�sak.
    If Not Vizsgalat Then
        ' Form bez�r�sa.
        Unload Me
    End If

    ' Akkor fut le ha az ideiglenes p�lya nem egyezne meg a konfigban t�rolt p�lya nev�vel.
    If Not tname = Config.Globalis_PalyaNeve Then
        ' K�r�k sz�m�nak be�ll�t�sa.
        Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
        ' Konfig f�jl be�ll�t�sa.
        Config.SetConfig

        ' Kezd�elem be�ll�t�sa.
        KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
        ' Megv�ltoztatjuk a k�r�k sz�m�nak ki�r�st.
        Palya.SetKorokSzama Palya.GetKezdokorErteke
        ' �j aut�k l�trehoz�sa.
        Palya.UjAutokLetrehozasa PalyaInfo.AutokSzama
    End If

    ' Friss�t�si az aut�k nyomvonal�nak megjelen�t�s�t.
    Palya.SetAutokNyomvonal
    ' Friss�t�si az aut�k t�k�letes k�r�z�s�nek �llapot�t.
    Palya.SetAutokTokeletesKorozes
    ' P�ly�hoz tartoz� k�r sz�m�nak ki�r�sa.
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    ' Friss�ti a virtu�lisan l�trehozott p�ly�t.
    Palya.VirtualisPalya_Frissites
End Sub

Private Sub CmdAlapertelmezes_Click()
    ' Konfig f�jl t�rl�se.
    Config.DeleteConfig
    ' Konfig f�jl bet�lt�se.
    Config.LoadConfig

    ' Alap�rtelmezett p�lya t�rl�se.
    Map.DeleteDefaultMap
    ' P�lya bet�lt�se.
    Map.LoadMap Config.Globalis_PalyaNeve
    ' K�r�k sz�m�nak be�ll�t�sa.
    Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
    ' Konfig f�jl be�ll�t�sa.
    Config.SetConfig

    ' Be�ll�t�sok inicializ�l�sa.
    Init
End Sub

' CheckNyomvonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckNyomvonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckNyomvonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_Nyomvonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckNyomvonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_Nyomvonal = False
    End If
End Sub

' CheckSzektorNevek gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckSzektorNevek_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckSzektorNevek.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_SzektorNevek = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckSzektorNevek.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_SzektorNevek = False
    End If
End Sub

' CheckStartCelVonalNeve gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckStartCelVonalNeve_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckStartCelVonalNeve.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_StartCelVonalNeve = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckStartCelVonalNeve.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_StartCelVonalNeve = False
    End If
End Sub

' CheckSzektorVonalak gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckSzektorVonalak_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckSzektorVonalak.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_SzektorVonalak = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckSzektorVonalak.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_SzektorVonalak = False
    End If
End Sub

' CheckStartCelVonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckStartCelVonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckStartCelVonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_StartCelVonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckStartCelVonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_StartCelVonal = False
    End If
End Sub

' CheckTokeletesKorozes gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckTokeletesKorozes_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckTokeletesKorozes.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_TokeletesKorozes = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckTokeletesKorozes.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempGlobalis_TokeletesKorozes = False
    End If
End Sub

' CheckElsoNyomvonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckElsoNyomvonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckElsoNyomvonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Elso_Nyomvonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckElsoNyomvonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Elso_Nyomvonal = False
    End If
End Sub

' CheckElsoTokeletesKorozes gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckElsoTokeletesKorozes_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckElsoTokeletesKorozes.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Elso_TokeletesKorozes = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckElsoTokeletesKorozes.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Elso_TokeletesKorozes = False
    End If
End Sub

' CheckMasodikNyomvonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckMasodikNyomvonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckMasodikNyomvonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Masodik_Nyomvonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckMasodikNyomvonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Masodik_Nyomvonal = False
    End If
End Sub

' CheckMasodikTokeletesKorozes gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckMasodikTokeletesKorozes_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckMasodikTokeletesKorozes.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Masodik_TokeletesKorozes = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckMasodikTokeletesKorozes.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Masodik_TokeletesKorozes = False
    End If
End Sub

' CheckHarmadikNyomvonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckHarmadikNyomvonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckHarmadikNyomvonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Harmadik_Nyomvonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckHarmadikNyomvonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Harmadik_Nyomvonal = False
    End If
End Sub

' CheckHarmadikTokeletesKorozes gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckHarmadikTokeletesKorozes_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckHarmadikTokeletesKorozes.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Harmadik_TokeletesKorozes = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckHarmadikTokeletesKorozes.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Harmadik_TokeletesKorozes = False
    End If
End Sub

' CheckNegyedikNyomvonal gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckNegyedikNyomvonal_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckNegyedikNyomvonal.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Negyedik_Nyomvonal = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckNegyedikNyomvonal.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Negyedik_Nyomvonal = False
    End If
End Sub

' CheckNegyedikTokeletesKorozes gomb esem�nye kattint�s hat�s�ra.
Private Sub CheckNegyedikTokeletesKorozes_Click()
    ' Ha egyenl� eggyel akkor fut le.
    If CheckNegyedikTokeletesKorozes.value = 1 Then
        ' Igazra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Negyedik_TokeletesKorozes = True
    ' Ha egyenl� null�val akkor fut le.
    ElseIf CheckNegyedikTokeletesKorozes.value = 0 Then
        ' Hamisra �ll�tja az ideiglenes v�ltoz�t.
        TempAutok_Negyedik_TokeletesKorozes = False
    End If
End Sub

' KorokComboBox lista esem�nye kattint�s hat�s�ra.
Private Sub KorokComboBox_Click()
    ' Ideiglenes k�r�k sz�m�nak megv�ltoztat�sa az index alapj�n.
    TempGlobalis_KorokSzama = CByte(Trim(KorokComboBox.List(KorokComboBox.ListIndex)))
End Sub

' PalyaComboBox lista esem�nye kattint�s hat�s�ra.
Private Sub PalyaComboBox_Click()
    ' Ideiglenes p�lya nev�nek megv�ltoztat�sa az index alapj�n.
    TempGlobalis_PalyaNeve = PalyaComboBox.List(PalyaComboBox.ListIndex)
End Sub

' Glob�lis list�n kiv�lasztjuk a megjelenitend� "frame"-t.
Private Sub GlobalisLista_Click()
    ' "Frame"-k l�thatatlann� t�tele.
    SetAllVisible False

    ' Kiv�lasztott elem alapj�n szelekt�lja melyik "frame" jelenjen meg.
    Select Case GlobalisLista.List(GlobalisLista.ListIndex)
        Case "�ltal�nos"
            ' Aut�k lista kijel�l�s�nek elt�ntet�se.
            AutokLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            Altalanos.Visible = True
        Case "P�lya"
            ' Aut�k lista kijel�l�s�nek elt�ntet�se.
            AutokLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            PalyaFrame.Visible = True
    End Select
End Sub

' Aut�k list�j�n kiv�lasztjuk a megjelenitend� "frame"-t.
Private Sub AutokLista_Click()
    ' "Frame"-k l�thatatlann� t�tele.
    SetAllVisible False

    ' Kiv�lasztott elem alapj�n szelekt�lja melyik "frame" jelenjen meg.
    Select Case AutokLista.List(AutokLista.ListIndex)
        Case "Els�"
            ' Glob�lis lista kijel�l�s�nek elt�ntet�se.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            ElsoFrame.Visible = True
        Case "M�sodik"
            ' Glob�lis lista kijel�l�s�nek elt�ntet�se.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            MasodikFrame.Visible = True
        Case "Harmadik"
            ' Glob�lis lista kijel�l�s�nek elt�ntet�se.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            HarmadikFrame.Visible = True
        Case "Negyedik"
            ' Glob�lis lista kijel�l�s�nek elt�ntet�se.
            GlobalisLista.ListIndex = -1
            ' "Frame" megjelen�t�se.
            NegyedikFrame.Visible = True
    End Select
End Sub

' Alap�rtelmezett folyamatok, be�ll�t�sok bet�lt�se.
Private Sub Init()
    ' P�lyaadatok be�ll�t�sa.
    SetPalyaComboBox

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        CheckNyomvonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    Else
        ' Kikapcsolja a pip�t.
        CheckNyomvonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_SzektorNevek Then
        ' Bekapcsolja a pip�t.
        CheckSzektorNevek.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    Else
        ' Kikapcsolja a pip�t.
        CheckSzektorNevek.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_StartCelVonalNeve Then
        ' Bekapcsolja a pip�t.
        CheckStartCelVonalNeve.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    Else
        ' Kikapcsolja a pip�t.
        CheckStartCelVonalNeve.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_SzektorVonalak Then
        ' Bekapcsolja a pip�t.
        CheckSzektorVonalak.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    Else
        ' Kikapcsolja a pip�t.
        CheckSzektorVonalak.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_StartCelVonal Then
        ' Bekapcsolja a pip�t.
        CheckStartCelVonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    Else
        ' Kikapcsolja a pip�t.
        CheckStartCelVonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Globalis_TokeletesKorozes Then
        ' Bekapcsolja a pip�t.
        CheckTokeletesKorozes.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    Else
        ' Kikapcsolja a pip�t.
        CheckTokeletesKorozes.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Elso_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        CheckElsoNyomvonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    Else
        ' Kikapcsolja a pip�t.
        CheckElsoNyomvonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Elso_TokeletesKorozes Then
        ' Bekapcsolja a pip�t.
        CheckElsoTokeletesKorozes.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    Else
        ' Kikapcsolja a pip�t.
        CheckElsoTokeletesKorozes.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Masodik_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        CheckMasodikNyomvonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    Else
        ' Kikapcsolja a pip�t.
        CheckMasodikNyomvonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Masodik_TokeletesKorozes Then
        ' Bekapcsolja a pip�t.
        CheckMasodikTokeletesKorozes.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    Else
        ' Kikapcsolja a pip�t.
        CheckMasodikTokeletesKorozes.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Harmadik_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        CheckHarmadikNyomvonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    Else
        ' Kikapcsolja a pip�t.
        CheckHarmadikNyomvonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Harmadik_TokeletesKorozes Then ' Kikapcsolva!!!
        ' Bekapcsolja a pip�t.
        CheckHarmadikTokeletesKorozes.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    Else
        ' Kikapcsolja a pip�t.
        CheckHarmadikTokeletesKorozes.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Negyedik_Nyomvonal Then
        ' Bekapcsolja a pip�t.
        CheckNegyedikNyomvonal.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    Else
        ' Kikapcsolja a pip�t.
        CheckNegyedikNyomvonal.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    End If

    ' Ha igaz az �rt�k akkor fut le.
    If Config.Autok_Negyedik_TokeletesKorozes Then
        ' Bekapcsolja a pip�t.
        CheckNegyedikTokeletesKorozes.value = 1
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    Else
        ' Kikapcsolja a pip�t.
        CheckNegyedikTokeletesKorozes.value = 0
        ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    End If

    ' Friss�t�si az aut�k nyomvonal�nak megjelen�t�s�t.
    Palya.SetAutokNyomvonal
    ' Friss�t�si az aut�k t�k�letes k�r�z�s�nek �llapot�t.
    Palya.SetAutokTokeletesKorozes
    ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
    TempGlobalis_PalyaNeve = Config.Globalis_PalyaNeve
    ' Kezd�elem be�ll�t�sa.
    KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
    ' Ideiglenes be�ll�t�s m�dos�t�sa a konfigban t�roltra.
    TempGlobalis_KorokSzama = Config.Globalis_KorokSzama
    ' Megv�ltoztatjuk a k�r�k sz�m�nak ki�r�st.
    Palya.SetKorokSzama Palya.GetKezdokorErteke
    ' P�ly�hoz tartoz� k�r sz�m�nak ki�r�sa.
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    ' Friss�ti a virtu�lisan l�trehozott p�ly�t.
    Palya.VirtualisPalya_Frissites
End Sub

' �sszes "frame" l�thatatlann� vagy l�that�v� t�tele.
' A "Visible" t�rolja a "frame"-k l�that�s�g�t.
Private Sub SetAllVisible(ByVal Visible As Boolean)
    ' L�that�s�g be�ll�t�sa.
    Altalanos.Visible = Visible
    ' L�that�s�g be�ll�t�sa.
    PalyaFrame.Visible = Visible
    ' L�that�s�g be�ll�t�sa.
    ElsoFrame.Visible = Visible
    ' L�that�s�g be�ll�t�sa.
    MasodikFrame.Visible = Visible
    ' L�that�s�g be�ll�t�sa.
    HarmadikFrame.Visible = Visible
    ' L�that�s�g be�ll�t�sa.
    NegyedikFrame.Visible = Visible
End Sub

Private Sub SetPalyaComboBox()
    ' K�nyvt�r adatok t�rol�sa.
    Dim a As String
    ' T�rolja p�lya nev�nek index�t.
    Dim index As Integer
    ' Mappa alk�nyvt�rainak �s f�jlainak bet�lt�se.
    a = Dir$(MapDir & "\*.*", vbDirectory)
    ' P�lyak sz�m�t null�za.
    PalyaInfo.PalyaNevekSzama = 0
    ' T�rlis a p�ly�k list�j�t.
    PalyaComboBox.Clear
    ' T�rlis a p�lya neveket a t�mbb�l.
    ReDim PalyaInfo.PalyaNevek(0 To 10) As String

    ' Addig fut am�g nem egyenl� az "a" a semmivel.
    Do While a <> ""
        ' Megn�zi hogy f�jl-e.
        If IsFile(MapDir & "\" & a) Then
            ' Ha nagyobb a p�ly�k sz�ma mint amit a t�mb t�rolni tud akkor fut le.
            If PalyaInfo.PalyaNevekSzama >= UBound(PalyaInfo.PalyaNevek) Then
                ' T�mb megn�vel�se.
                ReDim Preserve PalyaInfo.PalyaNevek(0 To UBound(PalyaInfo.PalyaNevek) + 10) As String
            End If

            ' Megn�zi hogy a konfigban t�rolt p�lya neve megegyezik-e a bet�lt�tt n�vvel.
            If a = Config.Globalis_PalyaNeve Then
                ' Elmenti a t�mbb�l az adott indexet.
                index = PalyaInfo.PalyaNevekSzama
            End If

            ' Ki�rja a p�lya nev�t.
            PalyaComboBox.AddItem a
            ' Elmenti a p�lya nev�t a t�mbbe.
            PalyaInfo.PalyaNevek(PalyaInfo.PalyaNevekSzama) = a
            ' Megn�veli a p�lya nevek sz�m�t.
            PalyaInfo.PalyaNevekSzama = PalyaInfo.PalyaNevekSzama + 1
        End If

        ' K�nyvt�r ugr�s.
        a = Dir$
    Loop

    ' Kezd�elem be�ll�t�sa.
    PalyaComboBox.ListIndex = index
End Sub
