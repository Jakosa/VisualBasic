VERSION 5.00
Begin VB.Form WarningForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Figyelmeztet�s!"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3390
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FigyelmeztetoJel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   120
      Picture         =   "WarningForm.frx":0000
      ScaleHeight     =   780
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Hiba 
      BackColor       =   &H8000000E&
      Caption         =   "Hiba"
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   435
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "WarningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fejl�c
' K�sz�tette: Bels� Vazul Istv�n
' Fejl�c v�ge

Option Explicit

' T�rolja a hiba�zenetet.
Public HibaUzenet As String
' Program le�ll�t�s�nak lehet�s�ge.
Public Leallitas As String

' Form bet�lt�se.
Private Sub Form_Load()
    ' Megvizsg�lja hogy a v�ltoz� igaz-e. Ha igen akkor �t�rja a hiba�zenetet.
    If Leallitas Then
        ' Le�ll�t�s �zenet hozz�ad�sa a hiba�zenethez.
        HibaUzenet = HibaUzenet & " A program az ok gombra kattint�s ut�n le fog �llni!"
    End If

    ' Hiba�zenet ki�r�sa.
    Hiba.Caption = HibaUzenet
    ' Hiba�zenet hossz�nak �tv�tele.
    Hiba.Width = TextWidth(HibaUzenet)

    ' Akkor fut le ha a "Hiba" hossza nagyobb mint a "FigyelmeztetoJel" + a "CmdOk" hossza. Ez az alap eset.
    If Hiba.Width > FigyelmeztetoJel.Width + CmdOk.Width Then
        Width = Hiba.Width + FigyelmeztetoJel.Width + CmdOk.Width / 2
    Else
        ' Akkor fut le ha nincs be�ll�tva hiba�zenetnek semmi se. �gy az ablak tov�bbra is �rtelmezhet� lesz �s nem mos�dik �ssze rajta az adat.
        Width = Hiba.Width + FigyelmeztetoJel.Width + CmdOk.Width
    End If

    ' Ablak k�z�prehelyez�se a k�perny�n.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' CmdOk gomb k�z�pre helyez�se a form-on bel�l.
    CmdOk.Left = Width / 2 - CmdOk.Width / 2
End Sub

' Form bez�r�sa.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Program leallitasanak esetleges lehet�s�g�t vizsgalja.
    Leallas

    ' Hiba�zenet t�rl�se.
    HibaUzenet = "Hiba"
End Sub

' CmdOk gomb esem�nye kattint�s hat�s�ra.
Private Sub CmdOk_Click()
    ' Program leallitasanak esetleges lehet�s�g�t vizsgalja.
    Leallas

    ' Hiba�zenet t�rl�se.
    HibaUzenet = "Hiba"
    ' Form bez�r�sa.
    Unload Me
End Sub

' Program leallitasanak esetleges lehet�s�g�t vizsgalja.
Private Sub Leallas()
    ' Megvizsg�lja hogy a v�ltoz� igaz-e. Ha igen le�ll a program.
    If Leallitas Then
        ' Program v�ge.
        End
    End If
End Sub

