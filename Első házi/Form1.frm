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
    Dim oldal(1 To 6) As Long                    ' T�rolja a dob�sokat oldalak szerint.
    Dim mgx As Long, jcs As Long                 ' Sz�mol�sra sz�lg�l� v�ltoz�k.

    For mgx = 1 To 6
        oldal(mgx) = 0                           ' Null�ra �ll�tja az oldalak dob�s�nak �rt�k�t.
    Next mgx

    Randomize
    Do While oldal(1) < 1000 And oldal(6) < 1000 ' Akkor l�p ki ha 1000-n�l nagyobb lesz az egyik �rt�k.
        jcs = Int(Rnd * 6 + 1)                   ' V�letlen sz�m 1-t�l 6-ig.
        oldal(jcs) = oldal(jcs) + 1              ' N�veli az adott oldal dob�s�nak �rt�k�t.
    Loop

    ListBox1.Clear                               ' T�rli a "ListBox"-ban t�rolt �rt�keket.

    For mgx = 1 To 6
        ListBox1.AddItem mgx & ". oldal dob�sainak sz�ma: " & oldal(mgx) ' V�gs� ki�r�s.
    Next mgx
End Sub
