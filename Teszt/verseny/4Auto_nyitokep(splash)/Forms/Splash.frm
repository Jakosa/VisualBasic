VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4043
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7005
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   120
         Top             =   3480
      End
      Begin VB.Image Image1 
         Height          =   1650
         Left            =   2760
         Picture         =   "Splash.frx":000C
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Verzió"
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
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
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
Option Explicit

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    Unload Me
'End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Verzió " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    Timer1.Enabled = True
End Sub

'Private Sub Frame1_Click()
'    Load_Palya
'    Unload Me
'End Sub

Private Sub Load_Palya()
    Palya.Show
End Sub

Private Sub Timer1_Timer()
    Load_Palya
    Unload Me
End Sub
