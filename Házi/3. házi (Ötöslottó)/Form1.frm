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
      Text            =   "Ide lesz ki�rva ha vesztett�l."
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Ide lesz ki�rva ha nyert�l."
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "90"
      ToolTipText     =   "�t�dik lehets�ges nyer�sz�munk! (Alap�rtelmez�sben: Az utols� sz�m az �t�slott�ban)"
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
      ToolTipText     =   "Harmadik lehets�ges nyer�sz�munk! (Alap�rtelmez�sben: A v�lasz az �letre, a vil�gmindens�gre, meg mindenre)"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "13"
      ToolTipText     =   "M�sodik lehets�ges nyer�sz�munk! (Alap�rtelmez�sben: P�ntek 13)"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "6"
      ToolTipText     =   "Els� lehets�ges nyer�sz�munk! (Alap�rtelmez�sben: Sz�let�sem napja)"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Nyerjunk 
      Caption         =   "5-�s lott� f�nyerem�ny?    Majd Kider�l!"
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
    Dim VeletlenOtosSzamok(1 To 5) As Long ' G�p �ltal gener�lt 5 sz�m.
    Dim OtosSzamok(1 To 5) As Long         ' �ltalalunk megadott 5 sz�m.
    Dim Allj As Boolean                    ' Meg�ll�tja a v�gtelen ciklust (ha meg kell �ll�tani).
    Dim Nyertel As Long                    ' Ha nyert�l akkor 5 lesz a benne t�rol sz�m �rt�ke.
    Dim random As Long                     ' V�letlen sz�mokat t�rol.
    Dim jcs As Long                        ' For ciklus v�ltoz�ja.
    Dim mgx As Long                        ' For ciklus v�ltoz�ja.

    OtosSzamok(1) = Text1.Text             ' Els� sz�m amit megadunk.
    OtosSzamok(2) = Text2.Text             ' M�sodik sz�m amit megadunk.
    OtosSzamok(3) = Text3.Text             ' Harmadik sz�m amit megadunk.
    OtosSzamok(4) = Text4.Text             ' Negyedik sz�m amit megadunk.
    OtosSzamok(5) = Text5.Text             ' �t�dik sz�m amit megadunk.

    ' Debug (Hibakeres�s)
    List1.Clear
    List1.AddItem "Hibakeres� ListBox"

    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        Allj = False

        If OtosSzamok(jcs) > 90 Then       ' Megn�zz�k hogy a sz�mok nagyobbak-e 90-n�l.
            List1.AddItem ""
            List1.AddItem "Nagyobb a " & jcs & ". mez�ben a sz�m 90-n�l!"
            List1.AddItem "Ilyen pedig nem lehets�ges a j�t�k szab�lyai szerint!"
            Exit Sub
        ElseIf OtosSzamok(jcs) < 1 Then    ' Megn�zz�k hogy a sz�mok kisebbek-e 1-n�l.
            List1.AddItem ""
            List1.AddItem "Kisebb a " & jcs & ". mez�ben a sz�m 1-n�l!"
            List1.AddItem "Ilyen pedig nem lehets�ges a j�t�k szab�lyai szerint!"
            Exit Sub
        End If

        ' Megn�zz�k van-e egyenl� sz�m a t�mbben.
        Do While Not Allj
            For mgx = LBound(OtosSzamok) To UBound(OtosSzamok)
                If OtosSzamok(jcs) = OtosSzamok(mgx) And Not jcs = mgx Then
                    List1.AddItem ""
                    List1.AddItem "Megegyeznek az �t�slott� sz�mok az " & jcs & ". �s " & mgx & ". mez�ben!"
                    List1.AddItem "Ilyen pedig nem lehets�ges a j�t�k szab�lyai szerint!"
                    Exit Sub
                End If
            Next mgx

            If mgx = UBound(OtosSzamok) + 1 Then
                Allj = True
            End If
        Loop
    Next jcs

    ' Debug (Hibakeres�s v�ge)

    For jcs = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
        Allj = False

        Do While Not Allj
            Randomize
            random = Int(Rnd * 90 + 1)     ' V�letlen sz�m 1-t�l 90-ig

            For mgx = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
                If VeletlenOtosSzamok(mgx) = random Then
                    Exit For               ' Ha azonos a v�letlen sz�m az egyik t�rolttal akkor �jrakezz�k a sz�m gener�l�s�t.
                End If
            Next mgx

            If mgx = UBound(VeletlenOtosSzamok) + 1 Then
                Allj = True                ' Ha nem azonos akkor le�ll�tjuk a while ciklust.
            End If
        Loop

        VeletlenOtosSzamok(jcs) = random   ' Let�roljuk a v�letlen sz�mot.
    Next jcs

    ' Debug (Hibakeres�s)
    List1.AddItem ""
    List1.AddItem "Felhaszn�l� �ltal megadott sz�mok:"

    ' Ki�rjuk az �ltalunk megadott sz�mokat.
    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        List1.AddItem jcs & ". sz�m: " & OtosSzamok(jcs)
    Next jcs

    List1.AddItem ""
    List1.AddItem "Program �ltal gener�lt sz�mok:"

    ' Ki�rjuk a g�p �ltal gener�lt sz�mokat.
    For jcs = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
        List1.AddItem jcs & ". sz�m: " & VeletlenOtosSzamok(jcs)
    Next jcs

    ' Debug (Hibakeres�s v�ge)

    ' Debug (Hibakeres�s)
    ' Csak akkor vedd ki a kommenteket jelz� ' jelet ha el akarod �rni hogy a sz�mok megegyezenek �s ki�rja hogy "Nyert�l!"
    ' Figyelem! Ezt csak hibakeres�skor akt�v�ld!
    ' VeletlenOtosSzamok(1) = Text1.Text
    ' VeletlenOtosSzamok(2) = Text5.Text
    ' VeletlenOtosSzamok(3) = Text3.Text
    ' VeletlenOtosSzamok(4) = Text2.Text
    ' VeletlenOtosSzamok(5) = Text4.Text
    ' Debug (Hibakeres�s v�ge)

    For jcs = LBound(OtosSzamok) To UBound(OtosSzamok)
        Allj = False

        Do While Not Allj
            ' Megn�zz�k van-e azonos sz�m a mi �ltalunk megadott sz�mokkal.
            For mgx = LBound(VeletlenOtosSzamok) To UBound(VeletlenOtosSzamok)
                If OtosSzamok(jcs) = VeletlenOtosSzamok(mgx) Then
                    Nyertel = Nyertel + 1 ' Ha van azonos akkor n�velj�k ezt a v�ltoz�t.
                    Allj = True           ' Le�ll�tjuk a while ciklust.
                    Exit For              ' Kil�p�nk a for ciklusb�l.
                End If
            Next mgx

            If mgx = UBound(VeletlenOtosSzamok) + 1 Then
                Allj = True               ' Le�ll�tjuk a while ciklust.
            End If
        Loop
    Next jcs

    If Nyertel = UBound(OtosSzamok) - LBound(OtosSzamok) + 1 Then ' Ha 5 akkor nyert�nk.
        Text6.Text = "Nyert�l!"           ' Nyertes sz�veg.
        Text7.Text = ""                   ' Mivel nem vesztett�nk �gy semmi lesz.
    Else
        Text6.Text = ""                   ' Mivel nem nyert�nk �gy semmi lesz.
        Text7.Text = "Nem nyert�l!"       ' Vesztes sz�veg.
    End If
End Sub

