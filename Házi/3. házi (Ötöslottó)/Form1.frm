VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   5295
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Ide lesz kiírva ha vesztettél."
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Ide lesz kiírva ha nyertél."
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "90"
      ToolTipText     =   "Ötödik lehetséges nyerõszámunk! (Alapértelmezésben: Az utolsó szám az ötöslottóban)"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "88"
      ToolTipText     =   $"Form1.frx":0000
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "42"
      ToolTipText     =   "Harmadik lehetséges nyerõszámunk! (Alapértelmezésben: A válasz az életre, a világmindenségre, meg mindenre)"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "13"
      ToolTipText     =   "Második lehetséges nyerõszámunk! (Alapértelmezésben: Péntek 13)"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "6"
      ToolTipText     =   "Elsõ lehetséges nyerõszámunk! (Alapértelmezésben: Születésem napja)"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Nyerjunk 
      Caption         =   "5-ös lottó fõnyeremény?    Majd Kiderül!"
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Nyerjunk_Click()
    Dim VeletlenOtosSzamok(1 To 5) As Long ' Gép által generált 5 szám.
    Dim OtosSzamok(1 To 5) As Long         ' Általalunk megadott 5 szám.
    Dim Allj As Boolean                    ' Megállítja a végtelen ciklust (ha meg kell állítani).
    Dim Nyertel As Long                    ' Ha nyertél akkor 5 lesz a benne tárol szám értéke.
    Dim random As Long                     ' Véletlen számokat tárol.
    Dim jcs As Long                        ' For ciklus változója.
    Dim mgx As Long                        ' For ciklus változója.

    OtosSzamok(1) = Text1.Text             ' Elsõ szám amit megadunk.
    OtosSzamok(2) = Text2.Text             ' Második szám amit megadunk.
    OtosSzamok(3) = Text3.Text             ' Harmadik szám amit megadunk.
    OtosSzamok(4) = Text4.Text             ' Negyedik szám amit megadunk.
    OtosSzamok(5) = Text5.Text             ' Ötödik szám amit megadunk.

    ' Debug (Hibakeresés)
    List1.Clear
    List1.AddItem "Hibakeresõ ListBox"

    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        Allj = False

        If OtosSzamok(jcs) > 90 Then       ' Megnézzük hogy a számok nagyobbak-e 90-nél.
            List1.AddItem ""
            List1.AddItem "Nagyobb a " & jcs & ". mezõben a szám 90-nél!"
            List1.AddItem "Ilyen pedig nem lehetséges a játék szabályai szerint!"
            Exit Sub
        ElseIf OtosSzamok(jcs) < 1 Then    ' Megnézzük hogy a számok kisebbek-e 1-nél.
            List1.AddItem ""
            List1.AddItem "Kisebb a " & jcs & ". mezõben a szám 1-nél!"
            List1.AddItem "Ilyen pedig nem lehetséges a játék szabályai szerint!"
            Exit Sub
        End If

        ' Megnézzük van-e egyenlõ szám a tömbben.
        Do While Not Allj
            For mgx = LBound(OtosSzamok) To UBound(OtosSzamok)
                If OtosSzamok(jcs) = OtosSzamok(mgx) And Not jcs = mgx Then
                    List1.AddItem ""
                    List1.AddItem "Megegyeznek az ötöslottó számok az " & jcs & ". és " & mgx & ". mezõben!"
                    List1.AddItem "Ilyen pedig nem lehetséges a játék szabályai szerint!"
                    Exit Sub
                End If
            Next mgx

            If mgx = UBound(OtosSzamok) + 1 Then
                Allj = True
            End If
        Loop
    Next jcs

    ' Debug (Hibakeresés vége)

    For jcs = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
        Allj = False

        Do While Not Allj
            Randomize
            random = Int(Rnd * 90 + 1)     ' Véletlen szám 1-tõl 90-ig

            For mgx = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
                If VeletlenOtosSzamok(mgx) = random Then
                    Exit For               ' Ha azonos a véletlen szám az egyik tárolttal akkor újrakezzük a szám generálását.
                End If
            Next mgx

            If mgx = UBound(VeletlenOtosSzamok) + 1 Then
                Allj = True                ' Ha nem azonos akkor leállítjuk a while ciklust.
            End If
        Loop

        VeletlenOtosSzamok(jcs) = random   ' Letároljuk a véletlen számot.
    Next jcs

    ' Debug (Hibakeresés)
    List1.AddItem ""
    List1.AddItem "Felhasználó által megadott számok:"

    ' Kiírjuk az általunk megadott számokat.
    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        List1.AddItem jcs & ". szám: " & OtosSzamok(jcs)
    Next jcs

    List1.AddItem ""
    List1.AddItem "Program által generált számok:"

    ' Kiírjuk a gép által generált számokat.
    For jcs = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
        List1.AddItem jcs & ". szám: " & VeletlenOtosSzamok(jcs)
    Next jcs

    ' Debug (Hibakeresés vége)

    ' Debug (Hibakeresés)
    ' Csak akkor vedd ki a kommenteket jelzõ ' jelet ha el akarod érni hogy a számok megegyezenek és kiírja hogy "Nyertél!"
    ' Figyelem! Ezt csak hibakereséskor aktíváld!
    ' VeletlenOtosSzamok(1) = Text1.Text
    ' VeletlenOtosSzamok(2) = Text5.Text
    ' VeletlenOtosSzamok(3) = Text3.Text
    ' VeletlenOtosSzamok(4) = Text2.Text
    ' VeletlenOtosSzamok(5) = Text4.Text
    ' Debug (Hibakeresés vége)

    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        Allj = False

        Do While Not Allj
            ' Megnézzük van-e azonos szám a mi általunk megadott számokkal.
            For mgx = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
                If OtosSzamok(jcs) = VeletlenOtosSzamok(mgx) Then
                    Nyertel = Nyertel + 1 ' Ha van azonos akkor növeljük ezt a változót.
                    Allj = True           ' Leállítjuk a while ciklust.
                    Exit For              ' Kilépünk a for ciklusból.
                End If
            Next mgx

            If mgx = UBound(VeletlenOtosSzamok) + 1 Then
                Allj = True               ' Leállítjuk a while ciklust.
            End If
        Loop
    Next jcs

    If Nyertel = UBound(OtosSzamok) - LBound(OtosSzamok) + 1 Then ' Ha 5 akkor nyertünk.
        Text6.Text = "Nyertél!"           ' Nyertes szöveg.
        Text7.Text = ""                   ' Mivel nem vesztettünk így semmi lesz.
    Else
        Text6.Text = ""                   ' Mivel nem nyertünk így semmi lesz.
        Text7.Text = "Nem nyertél!"       ' Vesztes szöveg.
    End If
End Sub

