VERSION 5.00
Begin VB.Form SettingsForm 
   BackColor       =   &H8000000E&
   Caption         =   "Beállítások"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
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
Option Explicit

Private TempGlobalis_Nyomvonal As Boolean
Private TempGlobalis_SzektorNevek As Boolean
Private TempGlobalis_StartCelVonalNeve As Boolean
Private TempGlobalis_KorokSzama As Byte
Private TempGlobalis_PalyaNeve As String
Private TempGlobalis_SzektorVonalak As Boolean
Private TempGlobalis_StartCelVonal As Boolean
Private TempGlobalis_TokeletesKorozes As Boolean
Private TempAutok_Elso_Nyomvonal As Boolean
Private TempAutok_Elso_TokeletesKorozes As Boolean
Private TempAutok_Masodik_Nyomvonal As Boolean
Private TempAutok_Masodik_TokeletesKorozes As Boolean
Private TempAutok_Harmadik_Nyomvonal As Boolean
Private TempAutok_Harmadik_TokeletesKorozes As Boolean
Private TempAutok_Negyedik_Nyomvonal As Boolean
Private TempAutok_Negyedik_TokeletesKorozes As Boolean

Private Sub Form_Load()
    Dim lStyle As Long

    lStyle = GetWindowLong(GlobalisLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(GlobalisLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(GlobalisLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    GlobalisLista.ListIndex = 0

    lStyle = GetWindowLong(AutokLista.hWnd, GWL_STYLE)
    lStyle = lStyle And (Not WS_BORDER)
    Call SetWindowLong(AutokLista.hWnd, GWL_STYLE, lStyle)
    Call SetWindowPos(AutokLista.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE)

    Init ' Beállítások inicializálása.
End Sub

Private Sub CmdOk_Click()
    CmdAlkalmaz_Click
    Unload Me
End Sub

Private Sub CmdMegse_Click()
    Unload Me
End Sub

Private Sub CmdAlkalmaz_Click()
    Dim tname As String
    tname = Config.Globalis_PalyaNeve

    Config.Globalis_Nyomvonal = TempGlobalis_Nyomvonal
    Config.Globalis_SzektorNevek = TempGlobalis_SzektorNevek
    Config.Globalis_StartCelVonalNeve = TempGlobalis_StartCelVonalNeve
    Config.Globalis_KorokSzama = TempGlobalis_KorokSzama
    Config.Globalis_PalyaNeve = TempGlobalis_PalyaNeve
    Config.Globalis_SzektorVonalak = TempGlobalis_SzektorVonalak
    Config.Globalis_StartCelVonal = TempGlobalis_StartCelVonal
    Config.Globalis_TokeletesKorozes = TempGlobalis_TokeletesKorozes
    Config.Autok_Elso_Nyomvonal = TempAutok_Elso_Nyomvonal
    Config.Autok_Elso_TokeletesKorozes = TempAutok_Elso_TokeletesKorozes
    Config.Autok_Masodik_Nyomvonal = TempAutok_Masodik_Nyomvonal
    Config.Autok_Masodik_TokeletesKorozes = TempAutok_Masodik_TokeletesKorozes
    Config.Autok_Harmadik_Nyomvonal = TempAutok_Harmadik_Nyomvonal
    Config.Autok_Harmadik_TokeletesKorozes = TempAutok_Harmadik_TokeletesKorozes
    Config.Autok_Negyedik_Nyomvonal = TempAutok_Negyedik_Nyomvonal
    Config.Autok_Negyedik_TokeletesKorozes = TempAutok_Negyedik_TokeletesKorozes
    Config.SetConfig

    Palya.SetKorokSzama Palya.GetKezdokorErteke
    Palya.UjAutokLetrehozasa Palya.GetAutokSzama

    If Config.Globalis_Nyomvonal Then
        Palya.Nyomvonal.Checked = 1
    Else
        Palya.Nyomvonal.Checked = 0
    End If

    If Config.Globalis_TokeletesKorozes Then
        Palya.Tokeletes_Korozes.Checked = 1
    Else
        Palya.Tokeletes_Korozes.Checked = 0
    End If

    Map.LoadMap Config.Globalis_PalyaNeve

    If Not tname = Config.Globalis_PalyaNeve Then
        Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
        Config.SetConfig

        KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
        Palya.SetKorokSzama Palya.GetKezdokorErteke
        Palya.UjAutokLetrehozasa Palya.GetAutokSzama
    End If

    Palya.SetAutokNyomvonal
    Palya.SetAutokTokeletesKorozes
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    Palya.HamisPalya_Frissites
End Sub

Private Sub CmdAlapertelmezes_Click()
    Config.DeleteConfig
    Config.LoadConfig

    Map.LoadMap Config.Globalis_PalyaNeve
    Config.Globalis_KorokSzama = PalyaInfo.KorokSzama
    Config.SetConfig

    Init ' Beállítások inicializálása.
End Sub

Private Sub CheckNyomvonal_Click()
    If CheckNyomvonal.value = 1 Then
        TempGlobalis_Nyomvonal = True
    ElseIf CheckNyomvonal.value = 0 Then
        TempGlobalis_Nyomvonal = False
    End If
End Sub

Private Sub CheckSzektorNevek_Click()
    If CheckSzektorNevek.value = 1 Then
        TempGlobalis_SzektorNevek = True
    ElseIf CheckSzektorNevek.value = 0 Then
        TempGlobalis_SzektorNevek = False
    End If
End Sub

Private Sub CheckStartCelVonalNeve_Click()
    If CheckStartCelVonalNeve.value = 1 Then
        TempGlobalis_StartCelVonalNeve = True
    ElseIf CheckStartCelVonalNeve.value = 0 Then
        TempGlobalis_StartCelVonalNeve = False
    End If
End Sub

Private Sub CheckSzektorVonalak_Click()
    If CheckSzektorVonalak.value = 1 Then
        TempGlobalis_SzektorVonalak = True
    ElseIf CheckSzektorVonalak.value = 0 Then
        TempGlobalis_SzektorVonalak = False
    End If
End Sub

Private Sub CheckStartCelVonal_Click()
    If CheckStartCelVonal.value = 1 Then
        TempGlobalis_StartCelVonal = True
    ElseIf CheckStartCelVonal.value = 0 Then
        TempGlobalis_StartCelVonal = False
    End If
End Sub

Private Sub CheckTokeletesKorozes_Click()
    If CheckTokeletesKorozes.value = 1 Then
        TempGlobalis_TokeletesKorozes = True
    ElseIf CheckTokeletesKorozes.value = 0 Then
        TempGlobalis_TokeletesKorozes = False
    End If
End Sub

Private Sub CheckElsoNyomvonal_Click()
    If CheckElsoNyomvonal.value = 1 Then
        TempAutok_Elso_Nyomvonal = True
    ElseIf CheckElsoNyomvonal.value = 0 Then
        TempAutok_Elso_Nyomvonal = False
    End If
End Sub

Private Sub CheckElsoTokeletesKorozes_Click()
    If CheckElsoTokeletesKorozes.value = 1 Then
        TempAutok_Elso_TokeletesKorozes = True
    ElseIf CheckElsoTokeletesKorozes.value = 0 Then
        TempAutok_Elso_TokeletesKorozes = False
    End If
End Sub

Private Sub CheckMasodikNyomvonal_Click()
    If CheckMasodikNyomvonal.value = 1 Then
        TempAutok_Masodik_Nyomvonal = True
    ElseIf CheckMasodikNyomvonal.value = 0 Then
        TempAutok_Masodik_Nyomvonal = False
    End If
End Sub

Private Sub CheckMasodikTokeletesKorozes_Click()
    If CheckMasodikTokeletesKorozes.value = 1 Then
        TempAutok_Masodik_TokeletesKorozes = True
    ElseIf CheckMasodikTokeletesKorozes.value = 0 Then
        TempAutok_Masodik_TokeletesKorozes = False
    End If
End Sub

Private Sub CheckHarmadikNyomvonal_Click()
    If CheckHarmadikNyomvonal.value = 1 Then
        TempAutok_Harmadik_Nyomvonal = True
    ElseIf CheckHarmadikNyomvonal.value = 0 Then
        TempAutok_Harmadik_Nyomvonal = False
    End If
End Sub

Private Sub CheckHarmadikTokeletesKorozes_Click()
    If CheckHarmadikTokeletesKorozes.value = 1 Then
        TempAutok_Harmadik_TokeletesKorozes = True
    ElseIf CheckHarmadikTokeletesKorozes.value = 0 Then
        TempAutok_Harmadik_TokeletesKorozes = False
    End If
End Sub

Private Sub CheckNegyedikNyomvonal_Click()
    If CheckNegyedikNyomvonal.value = 1 Then
        TempAutok_Negyedik_Nyomvonal = True
    ElseIf CheckNegyedikNyomvonal.value = 0 Then
        TempAutok_Negyedik_Nyomvonal = False
    End If
End Sub

Private Sub CheckNegyedikTokeletesKorozes_Click()
    If CheckNegyedikTokeletesKorozes.value = 1 Then
        TempAutok_Negyedik_TokeletesKorozes = True
    ElseIf CheckNegyedikTokeletesKorozes.value = 0 Then
        TempAutok_Negyedik_TokeletesKorozes = False
    End If
End Sub

Private Sub KorokComboBox_Click()
    TempGlobalis_KorokSzama = CByte(Trim(KorokComboBox.List(KorokComboBox.ListIndex)))
End Sub

Private Sub PalyaComboBox_Click()
    TempGlobalis_PalyaNeve = PalyaComboBox.List(PalyaComboBox.ListIndex)
End Sub

Private Sub GlobalisLista_Click()
    SetAllVisible False

    Select Case GlobalisLista.List(GlobalisLista.ListIndex)
        Case "Általános"
            AutokLista.ListIndex = -1
            Altalanos.visible = True
        Case "Pálya"
            AutokLista.ListIndex = -1
            PalyaFrame.visible = True
    End Select
End Sub

Private Sub AutokLista_Click()
    SetAllVisible False

    Select Case AutokLista.List(AutokLista.ListIndex)
        Case "Elsõ"
            GlobalisLista.ListIndex = -1
            ElsoFrame.visible = True
        Case "Második"
            GlobalisLista.ListIndex = -1
            MasodikFrame.visible = True
        Case "Harmadik"
            GlobalisLista.ListIndex = -1
            HarmadikFrame.visible = True
        Case "Negyedik"
            GlobalisLista.ListIndex = -1
            NegyedikFrame.visible = True
    End Select
End Sub

Private Sub Init()
    SetPalyaComboBox

    If Config.Globalis_Nyomvonal Then
        CheckNyomvonal.value = 1
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    Else
        CheckNyomvonal.value = 0
        TempGlobalis_Nyomvonal = Config.Globalis_Nyomvonal
    End If

    If Config.Globalis_SzektorNevek Then
        CheckSzektorNevek.value = 1
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    Else
        CheckSzektorNevek.value = 0
        TempGlobalis_SzektorNevek = Config.Globalis_SzektorNevek
    End If

    If Config.Globalis_StartCelVonalNeve Then
        CheckStartCelVonalNeve.value = 1
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    Else
        CheckStartCelVonalNeve.value = 0
        TempGlobalis_StartCelVonalNeve = Config.Globalis_StartCelVonalNeve
    End If

    If Config.Globalis_SzektorVonalak Then
        CheckSzektorVonalak.value = 1
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    Else
        CheckSzektorVonalak.value = 0
        TempGlobalis_SzektorVonalak = Config.Globalis_SzektorVonalak
    End If

    If Config.Globalis_StartCelVonal Then
        CheckStartCelVonal.value = 1
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    Else
        CheckStartCelVonal.value = 0
        TempGlobalis_StartCelVonal = Config.Globalis_StartCelVonal
    End If

    If Config.Globalis_TokeletesKorozes Then
        CheckTokeletesKorozes.value = 1
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    Else
        CheckTokeletesKorozes.value = 0
        TempGlobalis_TokeletesKorozes = Config.Globalis_TokeletesKorozes
    End If

    If Config.Autok_Elso_Nyomvonal Then
        CheckElsoNyomvonal.value = 1
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    Else
        CheckElsoNyomvonal.value = 0
        TempAutok_Elso_Nyomvonal = Config.Autok_Elso_Nyomvonal
    End If

    If Config.Autok_Elso_TokeletesKorozes Then
        CheckElsoTokeletesKorozes.value = 1
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    Else
        CheckElsoTokeletesKorozes.value = 0
        TempAutok_Elso_TokeletesKorozes = Config.Autok_Elso_TokeletesKorozes
    End If

    If Config.Autok_Masodik_Nyomvonal Then
        CheckMasodikNyomvonal.value = 1
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    Else
        CheckMasodikNyomvonal.value = 0
        TempAutok_Masodik_Nyomvonal = Config.Autok_Masodik_Nyomvonal
    End If

    If Config.Autok_Masodik_TokeletesKorozes Then
        CheckMasodikTokeletesKorozes.value = 1
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    Else
        CheckMasodikTokeletesKorozes.value = 0
        TempAutok_Masodik_TokeletesKorozes = Config.Autok_Masodik_TokeletesKorozes
    End If

    If Config.Autok_Harmadik_Nyomvonal Then
        CheckHarmadikNyomvonal.value = 1
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    Else
        CheckHarmadikNyomvonal.value = 0
        TempAutok_Harmadik_Nyomvonal = Config.Autok_Harmadik_Nyomvonal
    End If

    If Config.Autok_Harmadik_TokeletesKorozes Then
        CheckHarmadikTokeletesKorozes.value = 1
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    Else
        CheckHarmadikTokeletesKorozes.value = 0
        TempAutok_Harmadik_TokeletesKorozes = Config.Autok_Harmadik_TokeletesKorozes
    End If

    If Config.Autok_Negyedik_Nyomvonal Then
        CheckNegyedikNyomvonal.value = 1
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    Else
        CheckNegyedikNyomvonal.value = 0
        TempAutok_Negyedik_Nyomvonal = Config.Autok_Negyedik_Nyomvonal
    End If

    If Config.Autok_Negyedik_TokeletesKorozes Then
        CheckNegyedikTokeletesKorozes.value = 1
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    Else
        CheckNegyedikTokeletesKorozes.value = 0
        TempAutok_Negyedik_TokeletesKorozes = Config.Autok_Negyedik_TokeletesKorozes
    End If

    Palya.SetAutokNyomvonal
    Palya.SetAutokTokeletesKorozes
    TempGlobalis_PalyaNeve = Config.Globalis_PalyaNeve
    KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
    TempGlobalis_KorokSzama = Config.Globalis_KorokSzama
    Palya.SetKorokSzama Palya.GetKezdokorErteke
    PTKorokSzama.Caption = PalyaInfo.KorokSzama
    Palya.HamisPalya_Frissites
End Sub

Private Sub SetAllVisible(visible As Boolean)
    Altalanos.visible = visible
    PalyaFrame.visible = visible
    ElsoFrame.visible = visible
    MasodikFrame.visible = visible
    HarmadikFrame.visible = visible
    NegyedikFrame.visible = visible
End Sub

Private Sub SetPalyaComboBox()
    Dim a As String, index As Integer
    a = Dir$(MapDir & "\*.*", vbDirectory)
    PalyaInfo.PalyaNevekSzama = 0
    PalyaComboBox.Clear
    ReDim PalyaInfo.PalyaNevek(0 To 10) As String

    Do While a <> ""
        If IsFile(MapDir & "\" & a) Then
            If PalyaInfo.PalyaNevekSzama >= UBound(PalyaInfo.PalyaNevek) Then
                ReDim Preserve PalyaInfo.PalyaNevek(0 To UBound(PalyaInfo.PalyaNevek) + 10) As String
            End If

            If a = Config.Globalis_PalyaNeve Then
                index = PalyaInfo.PalyaNevekSzama
            End If

            PalyaComboBox.AddItem a
            PalyaInfo.PalyaNevek(PalyaInfo.PalyaNevekSzama) = a
            PalyaInfo.PalyaNevekSzama = PalyaInfo.PalyaNevekSzama + 1
        End If

        a = Dir$
    Loop

    PalyaComboBox.ListIndex = index
End Sub
