VERSION 5.00
Begin VB.Form WarningNewGame 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Figyelmeztet�s: �j j�t�k"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "M�gse"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton NoButton 
      Caption         =   "Nem"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton YesButton 
      Caption         =   "Igen"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "befejezett j�t�k v�geredm�ny�t. K�v�nja elmenteni?"
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
      BackColor       =   &H8000000E&
      Caption         =   "Figyelem! �j j�t�kot szeretne ind�tani de m�g nem mentette el az el�z�leg"
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
Option Explicit

Private Sub YesButton_Click()
    NewGameEnabled = False
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
