VERSION 5.00
Begin VB.Form WarningForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
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
   Begin VB.PictureBox WarningPicture 
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
Option Explicit

Public HibaUzenet As String

Private Sub Form_Load()
    Hiba.Caption = HibaUzenet
    Hiba.Width = TextWidth(HibaUzenet)

    If Hiba.Width > WarningPicture.Width + CmdOk.Width Then
        Width = Hiba.Width + WarningPicture.Width + CmdOk.Width / 2
    Else
        Width = Hiba.Width + WarningPicture.Width + CmdOk.Width
    End If

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    CmdOk.Left = Width / 2 - CmdOk.Width / 2
End Sub

Private Sub CmdOk_Click()
    HibaUzenet = "Hiba"
    Unload Me
End Sub

