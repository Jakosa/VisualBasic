VERSION 5.00
Begin VB.Form VForm 
   BackColor       =   &H8000000E&
   Caption         =   "Végeredmény"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   13155
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
   Begin VB.Label Label5 
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
Dim WithEvents Timer_Sorrend As VB.Timer
Attribute Timer_Sorrend.VB_VarHelpID = -1

Private Sub Form_Load()
    Label5.Caption = Palya.GetMKorokSzama & Label5.Caption

    Set Timer_Sorrend = VForm.Controls.Add("VB.Timer", "Timer_Sorrend", VForm)
    Timer_Sorrend.Interval = 500
End Sub

Private Sub Form_Terminate()
    Set Timer_Sorrend = Nothing
End Sub

Private Sub Timer_Sorrend_Timer()
    Dim tempkor As Byte, tempautok As Byte, ciklus As Integer, ciklus2 As Integer, i As Byte
    Dim NowTime As Date

    If Palya.GetKorokSzama > Palya.GetMKorokSzama Then
        tempkor = Palya.GetKorokSzama - 1
    Else
        tempkor = Palya.GetKorokSzama
    End If

    tempautok = 0
    CleanSText
    CleanOIText
    CleanOUText
    CleanLJText
    CleanKText

    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To Palya.GetAutokSzama
                If SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    Exit For
                ElseIf SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= Palya.GetAutokSzama Then
                    AddSText i & ". Autó: " & SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(Autok) To Palya.GetAutokSzama
                        If Autok(ciklus2).GetColor = SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            AddLJText Autok(ciklus2).GetLegjobbKorido & " másodperc"
                            AddOIText Autok(ciklus2).GetOsszKorido & " másodperc"
                            AddOUText Autok(ciklus2).GetOsszesUt & " m"

                            If tempautok = 0 Then
                                NowTime = SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                AddKText 0
                            Else
                                AddKText "+" & Abs(DateDiff("s", SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " másodperc"
                            End If
                        End If
                    Next ciklus2

                    tempautok = tempautok + 1
                End If

                If tempautok = Palya.GetAutokSzama Then
                    Exit For
                End If
            Next i

            If tempautok = Palya.GetAutokSzama Then
                Exit For
            End If
        Next ciklus

        If tempautok = Palya.GetAutokSzama Then
            Exit Do
        End If

        If tempkor > Palya.GetKezdokorErteke Then
            tempkor = tempkor - 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub CleanSText()
    SorrendText.Text = ""
End Sub

Private Sub AddSText(Szoveg As String)
    SorrendText.Text = SorrendText.Text & Szoveg & vbCrLf
End Sub

Private Sub CleanOUText()
    OsszUtText.Text = ""
End Sub

Private Sub AddOUText(Szoveg As String)
    OsszUtText.Text = OsszUtText.Text & Szoveg & vbCrLf
End Sub

Private Sub CleanOIText()
    OsszIdoText.Text = ""
End Sub

Private Sub AddOIText(Szoveg As String)
    OsszIdoText.Text = OsszIdoText.Text & Szoveg & vbCrLf
End Sub

Private Sub CleanLJText()
    LegjobbIdoText.Text = ""
End Sub

Private Sub AddLJText(Szoveg As String)
    LegjobbIdoText.Text = LegjobbIdoText.Text & Szoveg & vbCrLf
End Sub

Private Sub CleanKText()
    KulonbsegText.Text = ""
End Sub

Private Sub AddKText(Szoveg As String)
    KulonbsegText.Text = KulonbsegText.Text & Szoveg & vbCrLf
End Sub
