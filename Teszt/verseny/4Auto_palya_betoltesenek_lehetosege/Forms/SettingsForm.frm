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
      TabIndex        =   15
      Top             =   120
      Width           =   9375
      Begin VB.ComboBox PalyaComboBox 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1740
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
         TabIndex        =   11
         Top             =   1440
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
         Top             =   1440
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
         Sorted          =   -1  'True
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

    Dim a As String, index As Integer
    a = Dir$(MapDir & "\*.*", vbDirectory)
    PalyaInfo.PalyaNevekSzama = 0
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
    Config.Globalis_Nyomvonal = TempGlobalis_Nyomvonal
    Config.Globalis_SzektorNevek = TempGlobalis_SzektorNevek
    Config.Globalis_StartCelVonalNeve = TempGlobalis_StartCelVonalNeve
    Config.Globalis_KorokSzama = TempGlobalis_KorokSzama
    Config.Globalis_PalyaNeve = TempGlobalis_PalyaNeve
    Config.SetConfig

    Palya.SetKorokSzama Palya.GetKezdokorErteke
    Palya.TempAutoLista = -1

    If Config.Globalis_Nyomvonal Then
        Palya.Nyomvonal.Checked = 1
    Else
        Palya.Nyomvonal.Checked = 0
    End If

    Map.LoadMap Config.Globalis_PalyaNeve
    Palya.SetAutokNyomvonal
    Palya.SetSzektorNevek
    PalyaInfo.StartCelVonalNev.Label.visible = Config.Globalis_StartCelVonalNeve
    Palya.HamisPalya_Frissites
End Sub

Private Sub CmdAlapertelmezes_Click()
    Config.DeleteConfig
    Config.LoadConfig
    Map.LoadMap Config.Globalis_PalyaNeve
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
        Case "Második"
            GlobalisLista.ListIndex = -1
        Case "Harmadik"
            GlobalisLista.ListIndex = -1
        Case "Negyedik"
            GlobalisLista.ListIndex = -1
    End Select
End Sub

Private Sub Init()
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

    TempGlobalis_PalyaNeve = Config.Globalis_PalyaNeve
    KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
    TempGlobalis_KorokSzama = Config.Globalis_KorokSzama
    Palya.SetKorokSzama Palya.GetKezdokorErteke

    Palya.SetSzektorNevek
    PalyaInfo.StartCelVonalNev.Label.visible = Config.Globalis_StartCelVonalNeve
    Palya.HamisPalya_Frissites
End Sub

Private Sub SetAllVisible(visible As Boolean)
    Altalanos.visible = visible
    PalyaFrame.visible = visible
End Sub
