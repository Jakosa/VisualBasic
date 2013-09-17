VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   9480
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.ListBox ListBox1 
      Height          =   6495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim oldal(1 To 6) As Long                    ' Tárolja a dobásokat oldalak szerint.
    Dim mgx As Long, jcs As Long                 ' Számolásra szólgáló változók.

    For mgx = 1 To 6
        oldal(mgx) = 0                           ' Nullára állítja az oldalak dobásának értékét.
    Next mgx

    Randomize
    Do While oldal(1) < 1000 And oldal(6) < 1000 ' Akkor lép ki ha 1000-nél nagyobb lesz az egyik érték.
        jcs = Int(Rnd * 6 + 1)                   ' Véletlen szám 1-tõl 6-ig.
        oldal(jcs) = oldal(jcs) + 1              ' Növeli az adott oldal dobásának értékét.
    Loop

    ListBox1.Clear                               ' Törli a "ListBox"-ban tárolt értékeket.

    For mgx = 1 To 6
        ListBox1.AddItem mgx & ". oldal dobásainak száma: " & oldal(mgx) ' Végsõ kiírás.
    Next mgx
End Sub
