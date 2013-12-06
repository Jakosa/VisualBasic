VERSION 5.00
Begin VB.Form WarningForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Figyelmeztetés!"
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
' Fejléc
' Készítette: Belsõ Vazul István
' Fejléc vége

Option Explicit

' Tárolja a hibaüzenetet.
Public HibaUzenet As String
' Program leállításának lehetõsége.
Public Leallitas As String

' Form betöltése.
Private Sub Form_Load()
    ' Megvizsgálja hogy a változó igaz-e. Ha igen akkor átírja a hibaüzenetet.
    If Leallitas Then
        ' Leállítás üzenet hozzáadása a hibaüzenethez.
        HibaUzenet = HibaUzenet & " A program az ok gombra kattintás után le fog állni!"
    End If

    ' Hibaüzenet kiírása.
    Hiba.Caption = HibaUzenet
    ' Hibaüzenet hosszának átvétele.
    Hiba.Width = TextWidth(HibaUzenet)

    ' Akkor fut le ha a "Hiba" hossza nagyobb mint a "FigyelmeztetoJel" + a "CmdOk" hossza. Ez az alap eset.
    If Hiba.Width > FigyelmeztetoJel.Width + CmdOk.Width Then
        Width = Hiba.Width + FigyelmeztetoJel.Width + CmdOk.Width / 2
    Else
        ' Akkor fut le ha nincs beállítva hibaüzenetnek semmi se. Így az ablak továbbra is értelmezhetõ lesz és nem mosódik össze rajta az adat.
        Width = Hiba.Width + FigyelmeztetoJel.Width + CmdOk.Width
    End If

    ' Ablak középrehelyezése a képernyön.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' CmdOk gomb középre helyezése a form-on belül.
    CmdOk.Left = Width / 2 - CmdOk.Width / 2
End Sub

' Form bezárása.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Program leallitasanak esetleges lehetõségét vizsgalja.
    Leallas

    ' Hibaüzenet törlése.
    HibaUzenet = "Hiba"
End Sub

' CmdOk gomb eseménye kattintás hatására.
Private Sub CmdOk_Click()
    ' Program leallitasanak esetleges lehetõségét vizsgalja.
    Leallas

    ' Hibaüzenet törlése.
    HibaUzenet = "Hiba"
    ' Form bezárása.
    Unload Me
End Sub

' Program leallitasanak esetleges lehetõségét vizsgalja.
Private Sub Leallas()
    ' Megvizsgálja hogy a változó igaz-e. Ha igen leáll a program.
    If Leallitas Then
        ' Program vége.
        End
    End If
End Sub

