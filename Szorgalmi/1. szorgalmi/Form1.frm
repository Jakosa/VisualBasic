VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2175
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.ListBox ListBox1 
      Height          =   10395
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim dobasok(100 To 600) As Long                 ' T�rolja a dob�soknak az �sszeg�t.
    Dim mgx As Long, jcs As Long, osszeg As Long    ' Sz�mol�sra sz�lg�l� v�ltoz�k.

    For mgx = 100 To 600
        dobasok(mgx) = 0                            ' Null�ra �ll�tja a dob�sok �rt�k�t.
    Next mgx

    Randomize
    For jcs = 1 To 10000
        osszeg = 0

        For mgx = 1 To 100
            If osszeg = 0 Then                      ' Csak szebben mutat hogy nem null�hoz adja az�rt van felt�tel.
                osszeg = Int(Rnd * 6 + 1)           ' V�letlen sz�m 1-t�l 6-ig.
            Else
                osszeg = osszeg + Int(Rnd * 6 + 1)  ' V�letlen sz�m 1-t�l 6-ig. Hozz�adja az el�z�h�z.
            End If
        Next mgx

        dobasok(osszeg) = dobasok(osszeg) + 1
    Next jcs

    ListBox1.Clear                                  ' T�rli a "ListBox"-ban t�rolt �rt�keket.

    For mgx = 100 To 600
        ListBox1.AddItem "Dob�sok �sszege: " & mgx & " Dob�sok sz�ma: " & dobasok(mgx) ' V�gs� ki�r�s.
    Next mgx
End Sub
