VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   7200
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim T(1 To 15) As Long
    Dim i As Long, n As Long, osszeg As Long, atlag As Double

    List1.Clear                                  ' Takar�tjuk a listbox-ot.
    n = 0                                        ' Null�ra �ll�tjuk.
    osszeg = 0                                   ' Null�ra �ll�tjuk.

    Randomize
    For i = LBound(T) To UBound(T)
        T(i) = Int(Rnd * 101)                    ' V�letlen sz�m 0-t�l 100-ig.
        List1.AddItem i & ".elem: " & T(i)       ' Ki�rjuk azt hogy h�nyadik elem �s mi az elem �rt�ke.
        osszeg = osszeg + T(i)                   ' �sszeadjuk a sz�mokat.
    Next i

    atlag = osszeg / UBound(T) - LBound(T) + 1   ' �tlagoljuk a sz�mok �sszeg�t.
    List1.AddItem "�tlag: " & Round(atlag, 2)    ' Ki�rjuk az �tlagot k�t t�zedes pontos�ggal.
    List1.AddItem ""                             ' Sort�r�s.
    List1.AddItem "�tlag feletti:"

    For i = LBound(T) To UBound(T)
        If T(i) >= atlag Then                    ' Szelekt�ljuk az �tlag feletieket.
            n = n + 1                            ' Megsz�moljuk h�ny �tlag feletti van.
        End If
    Next i

    ReDim TT(1 To n) As Long                     ' L�trehozunk egy �j t�mb�t aminek a fent t�rol elemsz�m lesz az intervaluma. 1 -t�l n-ig.
    n = 1                                        ' Egyre �ll�tjuk az�rt mert a t�mb els� eleme lesz.

    For i = LBound(T) To UBound(T)
        If T(i) >= atlag Then                    ' Szelekt�ljuk az �tlag feletieket.
            TT(n) = T(i)                         ' Felt�ltj�k az �j t�mb�t az �tlag feletti elemekkel.
            n = n + 1                            ' N�velj�k az �rt�ket hogy ugorjunk a k�vetkez� elem�re a t�mbnek.
        End If
    Next i

    osszeg = 0                                   ' Null�ra �ll�tjuk.

    For i = LBound(TT) To UBound(TT)
        List1.AddItem i & ".elem: " & TT(i)      ' Ki�rjuk azt hogy h�nyadik elem �s mi az elem �rt�ke.
        osszeg = osszeg + TT(i)                  ' �sszeadjuk a sz�mokat.
    Next i

    atlag = osszeg / UBound(TT) - LBound(TT) + 1 ' �tlagoljuk a sz�mok �sszeg�t.
    List1.AddItem "�tlag: " & Round(atlag, 2)    ' Ki�rjuk az �tlagot k�t t�zedes pontos�ggal.
End Sub
