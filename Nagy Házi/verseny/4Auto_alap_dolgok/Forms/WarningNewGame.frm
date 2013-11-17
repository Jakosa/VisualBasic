VERSION 5.00
Begin VB.Form WarningNewGame 
   Caption         =   "Figyelmeztetés: Új játék"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Caption         =   "Mégse"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton NoButton 
      Caption         =   "Nem"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton YesButton 
      Caption         =   "Igen"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "befejezett játék végeredményét. Kívánja elmenteni?"
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
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   6405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Figyelem! Új játékot szeretne indítani de még nem mentette el az elõzõleg"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9195
   End
End
Attribute VB_Name = "WarningNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    NewGameEnabled = False
    Palya.NewGame_Click
    Unload Me
End Sub

Private Sub NoButton_Click()
    NewGameEnabled = True
    Palya.NewGame_Click
    Unload Me
End Sub

Private Sub YesButton_Click()
    NewGameEnabled = False
    Palya.NewGame_Click
    Unload Me
End Sub
