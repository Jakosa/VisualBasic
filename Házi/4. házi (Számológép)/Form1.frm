VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   DrawWidth       =   10
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton gyok 
      Caption         =   "sqrt"
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      ToolTipText     =   "N�gyzetgy�kvon�s"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton negyzet 
      Caption         =   "^2"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      ToolTipText     =   "N�gyzetre emel�s"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton osztas 
      Caption         =   "/"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      ToolTipText     =   "Oszt�s"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton szorzas 
      Caption         =   "*"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Szorz�s"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton minusz 
      Caption         =   "-"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Kivon�s"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton plusz 
      Caption         =   "+"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "�sszead�s"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox vegeredmeny 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      ToolTipText     =   "Ki�rja a v�geredm�nyt vagy a hib�t ha van."
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox bemenet2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "M�sodik bemenet. Harv�nyoz�sn�l �s gy�kvon�sn�l nincs haszn�lva."
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox bemenet1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Els� bemenet. Mindig haszn�latban van."
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sz�mol�g�p"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3390
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' �sszead�s
Private Sub plusz_Click()
    If Hibakereses(False) Then                                    ' Hibakeres�s. A m�sodik bemenet bekapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = CDbl(bemenet1.Text) + CDbl(bemenet2.Text)  ' �sszeadja a k�t sz�mot amit megadunk.
End Sub

' Kivon�s
Private Sub minusz_Click()
    If Hibakereses(False) Then                                    ' Hibakeres�s. A m�sodik bemenet bekapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = CDbl(bemenet1.Text) - CDbl(bemenet2.Text)  ' Kivonja az �ltalunk megadott sz�mokat egym�sb�l.
End Sub

' Szorz�s
Private Sub szorzas_Click()
    If Hibakereses(False) Then                                    ' Hibakeres�s. A m�sodik bemenet bekapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = CDbl(bemenet1.Text) * CDbl(bemenet2.Text)  ' �sszeszorozza az �ltalunk megadott sz�mokat.
End Sub

' Oszt�s
Private Sub osztas_Click()
    If Hibakereses(False) Then                                    ' Hibakeres�s. A m�sodik bemenet bekapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = CDbl(bemenet1.Text) / CDbl(bemenet2.Text)  ' A megadott sz�mokat elossza egym�ssal.
End Sub

' N�gyzetre emel�s
Private Sub negyzet_Click()
    If Hibakereses(True) Then                                     ' Hibakeres�s. A m�sodik bemenet kikapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    bemenet2.Text = ""                                            ' M�sodik bemenetet null�za mert nincs haszn�lva.
    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = CDbl(bemenet1.Text) ^ 2                    ' Az els� bemenetben l�v� sz�mot n�gyzetre emeli.
End Sub

' N�gyzetgy�kvon�s
Private Sub gyok_Click()
    If Hibakereses(True) Then                                     ' Hibakeres�s. A m�sodik bemenet kikapcsolva.
        Exit Sub                                                  ' Kil�p�s a folyamatb�l hiba miatt.
    End If

    bemenet2.Text = ""                                            ' M�sodik bemenetet null�za mert nincs haszn�lva.
    vegeredmeny.Alignment = 1                                     ' Sz�veg jobbra z�r�sa.
    vegeredmeny.Text = Sqr(CDbl(bemenet1.Text))                   ' Az els� bemenetben l�v� sz�mb�l n�gyzetgy�k�t von.
End Sub

' Leelen�rzi hogy van-e a programba valamilyen hiba. Ha igen ki�rja.
' Van egy v�ltoz� a "nincsbemenet2" amivel kikapcsolhat� a m�sodik bemenet hibakeres�se �gy figyelmen k�v�l hagyhat� a benne l�v� adat.
' Erre a n�gyzetre emel�s �s n�gyzetgy�kvon�s miatt van sz�ks�g.
Private Function Hibakereses(nincsbemenet2 As Boolean) As Boolean
    ' Megn�zz�k �res-e az els� bemenet. Ha igen akkor hiba van.
    If bemenet1.Text = "" Then
        vegeredmeny.Alignment = 0                                 ' Sz�veg balra z�r�sa.
        vegeredmeny.Text = "Hiba! Nincs �rt�k megadva az els� bemenetben!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilp�s a hibakeres�sb�l.
    End If

    ' Megn�zz�k �res-e a m�sodik bemenet. Ha igen akkor hiba van.
    ' Tov�bb� ha n�gyzetre emel�s vagy gy�kvon�s van akkor nem fut le ez a felt�tel.
    If bemenet2.Text = "" And Not nincsbemenet2 Then
        vegeredmeny.Alignment = 0                                 ' Sz�veg balra z�r�sa.
        vegeredmeny.Text = "Hiba! Nincs �rt�k megadva a m�sodik bemenetben!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilp�s a hibakeres�sb�l.
    End If

    ' Megn�zz�k sz�m-e az els� bemenet. Ha nem akkor hiba van.
    If Not IsNumeric(bemenet1.Text) Then
        vegeredmeny.Alignment = 0                                 ' Sz�veg balra z�r�sa.
        vegeredmeny.Text = "Hiba! Az els� bemenetbe nem sz�m lett megadva!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilp�s a hibakeres�sb�l.
    End If

    ' Megn�zz�k sz�m-e a m�sodik bemenet. Ha nem akkor hiba van.
    ' Tov�bb� ha n�gyzetre emel�s vagy gy�kvon�s van akkor nem fut le ez a felt�tel.
    If Not IsNumeric(bemenet2.Text) And Not nincsbemenet2 Then
        vegeredmeny.Alignment = 0                                 ' Sz�veg balra z�r�sa.
        vegeredmeny.Text = "Hiba! A m�sodik bemenetbe nem sz�m lett megadva!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilp�s a hibakeres�sb�l.
    End If

    Hibakereses = False                                           ' Nincs hiba.
End Function
