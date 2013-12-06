VERSION 5.00
Begin VB.Form AboutForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Névjegy"
   ClientHeight    =   2805
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7845
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1936.061
   ScaleMode       =   0  'User
   ScaleWidth      =   7366.86
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FejlesztokNevenekSavja 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   7935
      Begin VB.Label Fejlesztok 
         AutoSize        =   -1  'True
         Caption         =   "Fejlesztõk: Csaba, Péter, Valter, Vazul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.PictureBox Lo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1650
      Left            =   474
      Picture         =   "About.frx":0000
      ScaleHeight     =   1158.85
      ScaleMode       =   0  'User
      ScaleWidth      =   1158.85
      TabIndex        =   0
      Top             =   109
      Width           =   1650
   End
   Begin VB.Label Leiras 
      BackColor       =   &H8000000E&
      Caption         =   $"About.frx":BD54
      ForeColor       =   &H80000015&
      Height          =   915
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   5085
   End
   Begin VB.Label TermekNeve 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Termék Neve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   435
      Left            =   2625
      TabIndex        =   2
      Top             =   240
      Width           =   2385
   End
   Begin VB.Label Verzio 
      BackColor       =   &H8000000E&
      Caption         =   "Verzió"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   225
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   2565
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fejléc
' Készítette: Belsõ Vazul István
' Fejléc vége

Option Explicit

' Form betöltése.
Private Sub Form_Load()
    ' Form címének megváltoztatása a program nevére.
    Caption = App.Title & " névjegye"
    ' Kiírja a fejlesztõ(k) nevét.
    Fejlesztok.Caption = "Fejlesztõ(k): " & App.ProductName
    ' "Fejlesztok" középre helyezése a form-on belül.
    Fejlesztok.Left = Width / 2 - Fejlesztok.Width / 2
    ' Kiírja a program verzióját. pl: 0.0.1
    Verzio.Caption = "Verzió " & App.Major & "." & App.Minor & "." & App.Revision
    ' Kiírja a program nevét.
    TermekNeve.Caption = App.Title
End Sub

