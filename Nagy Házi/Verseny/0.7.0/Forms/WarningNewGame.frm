VERSION 5.00
Begin VB.Form WarningNewGame 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Figyelmeztetés: Új játék"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   120
      Picture         =   "WarningNewGame.frx":0000
      ScaleHeight     =   780
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton NoButton 
      Caption         =   "Nem"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton YesButton 
      Caption         =   "Igen"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
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
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   6405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   9195
   End
End
Attribute VB_Name = "WarningNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub YesButton_Click()
    VegeredmenyMentese.Save
    NewGameEnabled = True
    Palya.NewGame_Click
    Unload Me
End Sub

Private Sub NoButton_Click()
    NewGameEnabled = True
    Palya.NewGame_Click
    Unload Me
End Sub

Private Sub CancelButton_Click()
    NewGameEnabled = False
    Palya.NewGame_Click
    Unload Me
End Sub
