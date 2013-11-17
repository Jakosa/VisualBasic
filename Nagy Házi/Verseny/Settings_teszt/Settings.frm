VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H8000000E&
   Caption         =   "Beállítások"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
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
         ItemData        =   "Settings.frx":0000
         Left            =   1920
         List            =   "Settings.frx":003D
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1080
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
         Top             =   1080
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
         ItemData        =   "Settings.frx":0095
         Left            =   480
         List            =   "Settings.frx":00A5
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
         ItemData        =   "Settings.frx":00CC
         Left            =   480
         List            =   "Settings.frx":00D6
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
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'\\ Get Window Long Indexes...
Public Enum enGetWindowLong
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum

'\\ Window Style
Public Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
End Enum

Public Enum enSetWindowPos
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum
'\\ Set window ...
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As enSetWindowPos) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public TempGlobalis_Nyomvonal As Boolean
Public TempGlobalis_SzektorNevek As Boolean
Public TempGlobalis_KorokSzama As Byte

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

    'AutokLista.ListIndex = 0

    Config.LoadConfig
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
    Config.Globalis_KorokSzama = TempGlobalis_KorokSzama
    Config.SetConfig
End Sub

Private Sub CmdAlapertelmezes_Click()
    Config.DeleteConfig
    Config.LoadConfig
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

Private Sub KorokComboBox_Click()
    TempGlobalis_KorokSzama = CByte(Trim(KorokComboBox.List(KorokComboBox.ListIndex)))
End Sub

Private Sub GlobalisLista_Click()
    SetAllVisible False

    Select Case GlobalisLista.List(GlobalisLista.ListIndex)
        Case "Általános"
            AutokLista.ListIndex = -1
            Altalanos.visible = True
        Case "Pálya"
            AutokLista.ListIndex = -1
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

    KorokComboBox.ListIndex = Config.Globalis_KorokSzama - 2
    TempGlobalis_KorokSzama = Config.Globalis_KorokSzama
End Sub

Private Sub SetAllVisible(visible As Boolean)
    Altalanos.visible = visible
End Sub
