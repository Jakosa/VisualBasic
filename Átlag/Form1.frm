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

    List1.Clear                                  ' Takarítjuk a listbox-ot.
    n = 0                                        ' Nullára állítjuk.
    osszeg = 0                                   ' Nullára állítjuk.

    Randomize
    For i = LBound(T) To UBound(T)
        T(i) = Int(Rnd * 101)                    ' Véletlen szám 0-tól 100-ig.
        List1.AddItem i & ".elem: " & T(i)       ' Kiírjuk azt hogy hányadik elem és mi az elem értéke.
        osszeg = osszeg + T(i)                   ' Összeadjuk a számokat.
    Next i

    atlag = osszeg / UBound(T) - LBound(T) + 1   ' Átlagoljuk a számok összegét.
    List1.AddItem "Átlag: " & Round(atlag, 2)    ' Kiírjuk az átlagot két tízedes pontosággal.
    List1.AddItem ""                             ' Sortörés.
    List1.AddItem "Átlag feletti:"

    For i = LBound(T) To UBound(T)
        If T(i) >= atlag Then                    ' Szelektáljuk az átlag feletieket.
            n = n + 1                            ' Megszámoljuk hány átlag feletti van.
        End If
    Next i

    ReDim TT(1 To n) As Long                     ' Létrehozunk egy új tömböt aminek a fent tárol elemszám lesz az intervaluma. 1 -tõl n-ig.
    n = 1                                        ' Egyre állítjuk azért mert a tömb elsõ eleme lesz.

    For i = LBound(T) To UBound(T)
        If T(i) >= atlag Then                    ' Szelektáljuk az átlag feletieket.
            TT(n) = T(i)                         ' Feltöltjük az új tömböt az átlag feletti elemekkel.
            n = n + 1                            ' Növeljük az értéket hogy ugorjunk a következõ elemére a tömbnek.
        End If
    Next i

    osszeg = 0                                   ' Nullára állítjuk.

    For i = LBound(TT) To UBound(TT)
        List1.AddItem i & ".elem: " & TT(i)      ' Kiírjuk azt hogy hányadik elem és mi az elem értéke.
        osszeg = osszeg + TT(i)                  ' Összeadjuk a számokat.
    Next i

    atlag = osszeg / UBound(TT) - LBound(TT) + 1 ' Átlagoljuk a számok összegét.
    List1.AddItem "Átlag: " & Round(atlag, 2)    ' Kiírjuk az átlagot két tízedes pontosággal.
End Sub
