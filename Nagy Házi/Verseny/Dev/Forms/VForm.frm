VERSION 5.00
Begin VB.Form VForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Végeredmény"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox KulonbsegText 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox LegjobbIdoText 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox OsszIdoText 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox OsszUtText 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox SorrendText 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Különbség"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   4
      X1              =   10440
      X2              =   10440
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Label Cimke 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   " körös verseny végeredménye."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   12960
      X2              =   120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   7920
      X2              =   7920
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   5400
      X2              =   5400
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Legjobb köridõ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   6
      Top             =   960
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   2880
      X2              =   2880
      Y1              =   960
      Y2              =   3480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Versenyidõ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6120
      TabIndex        =   4
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Összes megtett út"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10680
      TabIndex        =   3
      Top             =   960
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Sorrend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   885
   End
End
Attribute VB_Name = "VForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Sorrendet frissíti.
Private WithEvents Timer_Sorrend As VB.Timer
Attribute Timer_Sorrend.VB_VarHelpID = -1

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Form_Load()
    ' Cimke beállítása.
    Cimke.Caption = Config.Globalis_KorokSzama & Cimke.Caption

    ' Sorrend timer létrehozása
    Set Timer_Sorrend = VForm.Controls.Add("VB.Timer", "Timer_Sorrend", VForm)
    ' Érték beállítása. 500 millisec
    Timer_Sorrend.Interval = 500
End Sub

' Form megszünésekor bizonyos dolgok megsemisítésre kerülnek.
Private Sub Form_Terminate()
    ' Nullázás
    Set Timer_Sorrend = Nothing
End Sub

Private Sub Timer_Sorrend_Timer()
    ' Ideiglenes köröket tárol.
    Dim tempkor As Byte
    ' Ideiglenes autók számát tárolja.
    Dim tempautok As Byte
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Integer
    ' "ciklus2" segédváltozó a ciklushoz.
    Dim ciklus2 As Integer
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    ' Tárolja a szektor idõt.
    Dim NowTime As Date

    ' Ha a Palya.GetKorokSzama nagyobb mint a maximális körök száma akkor fut le.
    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        ' Érték beállítása. Azért -1 mert a változó a játék végén +1-el nagyobbra lett megnövelve.
        tempkor = Palya.GetKorokSzama - 1
    Else
        ' Érték beállítása.
        tempkor = Palya.GetKorokSzama
    End If

    ' Nullázás.
    tempautok = 0
    ' TextBox takarítása.
    CleanSText
    ' TextBox takarítása.
    CleanOIText
    ' TextBox takarítása.
    CleanOUText
    ' TextBox takarítása.
    CleanLJText
    ' TextBox takarítása.
    CleanKText

    ' Végtelenségig futó ciklus
    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                ' Akkor fut le ha nincs szin beállítva (nincs autó) és a van adat is.
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    ' Kilépés a ciklusból.
                    Exit For
                ' Akkor fut le ha van adat és az ideiglenes autók száma kisebb vagy engyenlõ az AutokSzama-val.
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                    ' Szöveg kiírása.
                    AddSText i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                        ' Akkor fut le ha a kocsi szine egyenlõ a szektorhoz tartózó kocsi szinével.
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            ' Szöveg kiírása.
                            AddLJText PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " másodperc"
                            ' Szöveg kiírása.
                            AddOIText PalyaInfo.Autok(ciklus2).GetOsszKorido & " másodperc"
                            ' Szöveg kiírása.
                            AddOUText PalyaInfo.Autok(ciklus2).GetOsszesUt & " m"

                            ' Akkor fut le ha az ideiglenes autók száma nem nulla.
                            If tempautok = 0 Then
                                ' Menti a szektor idejét.
                                NowTime = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                ' Szöveg kiírása.
                                AddKText 0
                            Else
                                ' Szöveg kiírása.
                                AddKText "+" & Abs(DateDiff("s", PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " másodperc"
                            End If
                        End If
                    Next ciklus2

                    ' Megnöveljük 1-el az ideiglenes autók számát.
                    tempautok = tempautok + 1
                End If

                ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kilépés a ciklusból.
                    Exit For
                End If
            Next i

            ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
            If tempautok = PalyaInfo.AutokSzama Then
                ' Kilépés a ciklusból.
                Exit For
            End If
        Next ciklus

        ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
        If tempautok = PalyaInfo.AutokSzama Then
            ' Kilépés a ciklusból.
            Exit Do
        End If

        ' Akkor fut le ha az ideiglenes körök száma nagyobb mind a kezdõkör értéke.
        If tempkor > Palya.GetKezdokorErteke Then
            ' Az ideiglenes körök számát csökkentjük eggyel.
            tempkor = tempkor - 1
        Else
            ' Kilépés a ciklusból.
            Exit Do
        End If
    Loop
End Sub

' Sorrend TextBox adatainak törlése.
Private Sub CleanSText()
    ' Érték beállítása.
    SorrendText.Text = ""
End Sub

' Sorrend TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddSText(ByVal Szoveg As String)
    ' Érték beállítása.
    SorrendText.Text = SorrendText.Text & Szoveg & vbCrLf
End Sub

' OsszUt TextBox adatainak törlése.
Private Sub CleanOUText()
    ' Érték beállítása.
    OsszUtText.Text = ""
End Sub

' OsszUt TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddOUText(ByVal Szoveg As String)
    ' Érték beállítása.
    OsszUtText.Text = OsszUtText.Text & Szoveg & vbCrLf
End Sub

' OsszIdo TextBox adatainak törlése.
Private Sub CleanOIText()
    ' Érték beállítása.
    OsszIdoText.Text = ""
End Sub

' OsszIdo TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddOIText(ByVal Szoveg As String)
    ' Érték beállítása.
    OsszIdoText.Text = OsszIdoText.Text & Szoveg & vbCrLf
End Sub

' LegjobbIdo TextBox adatainak törlése.
Private Sub CleanLJText()
    ' Érték beállítása.
    LegjobbIdoText.Text = ""
End Sub

' LegjobbIdo TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddLJText(ByVal Szoveg As String)
    ' Érték beállítása.
    LegjobbIdoText.Text = LegjobbIdoText.Text & Szoveg & vbCrLf
End Sub

' Kulonbseg TextBox adatainak törlése.
Private Sub CleanKText()
    ' Érték beállítása.
    KulonbsegText.Text = ""
End Sub

' Kulonbseg TextBox-hoz szöveg hozzáadása (Új sorba).
' A "Szoveg" tárolja azon szöveget amely kiírása fog kerülni.
Private Sub AddKText(ByVal Szoveg As String)
    ' Érték beállítása.
    KulonbsegText.Text = KulonbsegText.Text & Szoveg & vbCrLf
End Sub
