VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4215
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4043
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7005
      Begin VB.Timer Kesleltetes 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   120
         Top             =   3480
      End
      Begin VB.Image Lo 
         Height          =   1650
         Left            =   2760
         Picture         =   "Splash.frx":143A
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Verzio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Verzi�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6150
         TabIndex        =   1
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label TermekNeve 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Term�k neve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2400
         TabIndex        =   2
         Top             =   2160
         Width           =   2430
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fejl�c
' K�sz�tette: Jakosa Csaba �rp�d
' Fejl�c v�ge

Option Explicit

' Form bet�lt�se.
Private Sub Form_Load()
    ' Ki�rja a program verzi�j�t. pl: 0.0.1
    Verzio.Caption = "Verzi� " & App.Major & "." & App.Minor & "." & App.Revision
    ' Ki�rja a program nev�t.
    TermekNeve.Caption = App.Title

    ' Konfig f�jl bet�lt�se.
    Config.LoadConfig

    ' P�lya bet�lt�se.
    Map.LoadMap Config.Globalis_PalyaNeve

    ' K�sleltett programind�t�s elind�tva.
    Kesleltetes.Enabled = True
End Sub

' Palya form megnyit�sa.
Private Sub Load_Palya()
    ' Megnyit�s.
    Palya.Show
End Sub

' K�sleltet�s id�zit� elj�r�sa.
Private Sub Kesleltetes_Timer()
    ' Palya form megnyitasa.
    Load_Palya

    ' K�sleltet�sre szolg�l� id�zit� le�ll�t�sa.
    Kesleltetes.Enabled = False

    ' Form bez�r�sa.
    Unload Me
End Sub
