VERSION 5.00
Begin VB.Form VForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "V�geredm�ny"
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
      Caption         =   "K�l�nbs�g"
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
      Caption         =   " k�r�s verseny v�geredm�nye."
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
      Caption         =   "Legjobb k�rid�"
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
      Caption         =   "Versenyid�"
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
      Caption         =   "�sszes megtett �t"
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

' Sorrendet friss�ti.
Private WithEvents Timer_Sorrend As VB.Timer
Attribute Timer_Sorrend.VB_VarHelpID = -1

' Be�ll�tjuk a form l�trehoz�sakor az alap folyamatokat.
Private Sub Form_Load()
    ' Cimke be�ll�t�sa.
    Cimke.Caption = Config.Globalis_KorokSzama & Cimke.Caption

    ' Sorrend timer l�trehoz�sa
    Set Timer_Sorrend = VForm.Controls.Add("VB.Timer", "Timer_Sorrend", VForm)
    ' �rt�k be�ll�t�sa. 500 millisec
    Timer_Sorrend.Interval = 500
End Sub

' Form megsz�n�sekor bizonyos dolgok megsemis�t�sre ker�lnek.
Private Sub Form_Terminate()
    ' Null�z�s
    Set Timer_Sorrend = Nothing
End Sub

Private Sub Timer_Sorrend_Timer()
    ' Ideiglenes k�r�ket t�rol.
    Dim tempkor As Byte
    ' Ideiglenes aut�k sz�m�t t�rolja.
    Dim tempautok As Byte
    ' "ciklus" seg�dv�ltoz� a ciklushoz.
    Dim ciklus As Integer
    ' "ciklus2" seg�dv�ltoz� a ciklushoz.
    Dim ciklus2 As Integer
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    ' T�rolja a szektor id�t.
    Dim NowTime As Date

    ' Ha a Palya.GetKorokSzama nagyobb mint a maxim�lis k�r�k sz�ma akkor fut le.
    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        ' �rt�k be�ll�t�sa. Az�rt -1 mert a v�ltoz� a j�t�k v�g�n +1-el nagyobbra lett megn�velve.
        tempkor = Palya.GetKorokSzama - 1
    Else
        ' �rt�k be�ll�t�sa.
        tempkor = Palya.GetKorokSzama
    End If

    ' Null�z�s.
    tempautok = 0
    ' TextBox takar�t�sa.
    CleanSText
    ' TextBox takar�t�sa.
    CleanOIText
    ' TextBox takar�t�sa.
    CleanOUText
    ' TextBox takar�t�sa.
    CleanLJText
    ' TextBox takar�t�sa.
    CleanKText

    ' V�gtelens�gig fut� ciklus
    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                ' Akkor fut le ha nincs szin be�ll�tva (nincs aut�) �s a van adat is.
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    ' Kil�p�s a ciklusb�l.
                    Exit For
                ' Akkor fut le ha van adat �s az ideiglenes aut�k sz�ma kisebb vagy engyenl� az AutokSzama-val.
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                    ' Sz�veg ki�r�sa.
                    AddSText i & ". Aut�: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                        ' Akkor fut le ha a kocsi szine egyenl� a szektorhoz tart�z� kocsi szin�vel.
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            ' Sz�veg ki�r�sa.
                            AddLJText PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " m�sodperc"
                            ' Sz�veg ki�r�sa.
                            AddOIText PalyaInfo.Autok(ciklus2).GetOsszKorido & " m�sodperc"
                            ' Sz�veg ki�r�sa.
                            AddOUText PalyaInfo.Autok(ciklus2).GetOsszesUt & " m"

                            ' Akkor fut le ha az ideiglenes aut�k sz�ma nem nulla.
                            If tempautok = 0 Then
                                ' Menti a szektor idej�t.
                                NowTime = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                ' Sz�veg ki�r�sa.
                                AddKText 0
                            Else
                                ' Sz�veg ki�r�sa.
                                AddKText "+" & Abs(DateDiff("s", PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " m�sodperc"
                            End If
                        End If
                    Next ciklus2

                    ' Megn�velj�k 1-el az ideiglenes aut�k sz�m�t.
                    tempautok = tempautok + 1
                End If

                ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kil�p�s a ciklusb�l.
                    Exit For
                End If
            Next i

            ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
            If tempautok = PalyaInfo.AutokSzama Then
                ' Kil�p�s a ciklusb�l.
                Exit For
            End If
        Next ciklus

        ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
        If tempautok = PalyaInfo.AutokSzama Then
            ' Kil�p�s a ciklusb�l.
            Exit Do
        End If

        ' Akkor fut le ha az ideiglenes k�r�k sz�ma nagyobb mind a kezd�k�r �rt�ke.
        If tempkor > Palya.GetKezdokorErteke Then
            ' Az ideiglenes k�r�k sz�m�t cs�kkentj�k eggyel.
            tempkor = tempkor - 1
        Else
            ' Kil�p�s a ciklusb�l.
            Exit Do
        End If
    Loop
End Sub

' Sorrend TextBox adatainak t�rl�se.
Private Sub CleanSText()
    ' �rt�k be�ll�t�sa.
    SorrendText.Text = ""
End Sub

' Sorrend TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddSText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    SorrendText.Text = SorrendText.Text & Szoveg & vbCrLf
End Sub

' OsszUt TextBox adatainak t�rl�se.
Private Sub CleanOUText()
    ' �rt�k be�ll�t�sa.
    OsszUtText.Text = ""
End Sub

' OsszUt TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddOUText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    OsszUtText.Text = OsszUtText.Text & Szoveg & vbCrLf
End Sub

' OsszIdo TextBox adatainak t�rl�se.
Private Sub CleanOIText()
    ' �rt�k be�ll�t�sa.
    OsszIdoText.Text = ""
End Sub

' OsszIdo TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddOIText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    OsszIdoText.Text = OsszIdoText.Text & Szoveg & vbCrLf
End Sub

' LegjobbIdo TextBox adatainak t�rl�se.
Private Sub CleanLJText()
    ' �rt�k be�ll�t�sa.
    LegjobbIdoText.Text = ""
End Sub

' LegjobbIdo TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddLJText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    LegjobbIdoText.Text = LegjobbIdoText.Text & Szoveg & vbCrLf
End Sub

' Kulonbseg TextBox adatainak t�rl�se.
Private Sub CleanKText()
    ' �rt�k be�ll�t�sa.
    KulonbsegText.Text = ""
End Sub

' Kulonbseg TextBox-hoz sz�veg hozz�ad�sa (�j sorba).
' A "Szoveg" t�rolja azon sz�veget amely ki�r�sa fog ker�lni.
Private Sub AddKText(ByVal Szoveg As String)
    ' �rt�k be�ll�t�sa.
    KulonbsegText.Text = KulonbsegText.Text & Szoveg & vbCrLf
End Sub
