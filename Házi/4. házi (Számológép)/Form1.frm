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
      ToolTipText     =   "Négyzetgyökvonás"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton negyzet 
      Caption         =   "^2"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      ToolTipText     =   "Négyzetre emelés"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton osztas 
      Caption         =   "/"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      ToolTipText     =   "Osztás"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton szorzas 
      Caption         =   "*"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Szorzás"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton minusz 
      Caption         =   "-"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Kivonás"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton plusz 
      Caption         =   "+"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Összeadás"
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
      ToolTipText     =   "Kiírja a végeredményt vagy a hibát ha van."
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
      ToolTipText     =   "Második bemenet. Harványozásnál és gyõkvonásnál nincs használva."
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
      ToolTipText     =   "Elsõ bemenet. Mindig használatban van."
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Számológép"
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
' Összeadás
Private Sub plusz_Click()
    If Hibakereses(False) Then                                    ' Hibakeresés. A második bemenet bekapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = CDbl(bemenet1.Text) + CDbl(bemenet2.Text)  ' Összeadja a két számot amit megadunk.
End Sub

' Kivonás
Private Sub minusz_Click()
    If Hibakereses(False) Then                                    ' Hibakeresés. A második bemenet bekapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = CDbl(bemenet1.Text) - CDbl(bemenet2.Text)  ' Kivonja az általunk megadott számokat egymásból.
End Sub

' Szorzás
Private Sub szorzas_Click()
    If Hibakereses(False) Then                                    ' Hibakeresés. A második bemenet bekapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = CDbl(bemenet1.Text) * CDbl(bemenet2.Text)  ' Összeszorozza az általunk megadott számokat.
End Sub

' Osztás
Private Sub osztas_Click()
    If Hibakereses(False) Then                                    ' Hibakeresés. A második bemenet bekapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = CDbl(bemenet1.Text) / CDbl(bemenet2.Text)  ' A megadott számokat elossza egymással.
End Sub

' Négyzetre emelés
Private Sub negyzet_Click()
    If Hibakereses(True) Then                                     ' Hibakeresés. A második bemenet kikapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    bemenet2.Text = ""                                            ' Második bemenetet nulláza mert nincs használva.
    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = CDbl(bemenet1.Text) ^ 2                    ' Az elsõ bemenetben lévõ számot négyzetre emeli.
End Sub

' Négyzetgyökvonás
Private Sub gyok_Click()
    If Hibakereses(True) Then                                     ' Hibakeresés. A második bemenet kikapcsolva.
        Exit Sub                                                  ' Kilépés a folyamatból hiba miatt.
    End If

    bemenet2.Text = ""                                            ' Második bemenetet nulláza mert nincs használva.
    vegeredmeny.Alignment = 1                                     ' Szöveg jobbra zárása.
    vegeredmeny.Text = Sqr(CDbl(bemenet1.Text))                   ' Az elsõ bemenetben lévõ számból négyzetgyököt von.
End Sub

' Leelenõrzi hogy van-e a programba valamilyen hiba. Ha igen kiírja.
' Van egy változó a "nincsbemenet2" amivel kikapcsolható a második bemenet hibakeresése így figyelmen kívül hagyható a benne lévõ adat.
' Erre a négyzetre emelés és négyzetgyökvonás miatt van szükség.
Private Function Hibakereses(nincsbemenet2 As Boolean) As Boolean
    ' Megnézzük üres-e az elsõ bemenet. Ha igen akkor hiba van.
    If bemenet1.Text = "" Then
        vegeredmeny.Alignment = 0                                 ' Szöveg balra zárása.
        vegeredmeny.Text = "Hiba! Nincs érték megadva az elsõ bemenetben!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilpés a hibakeresésbõl.
    End If

    ' Megnézzük üres-e a második bemenet. Ha igen akkor hiba van.
    ' Továbbá ha négyzetre emelés vagy gyökvonás van akkor nem fut le ez a feltétel.
    If bemenet2.Text = "" And Not nincsbemenet2 Then
        vegeredmeny.Alignment = 0                                 ' Szöveg balra zárása.
        vegeredmeny.Text = "Hiba! Nincs érték megadva a második bemenetben!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilpés a hibakeresésbõl.
    End If

    ' Megnézzük szám-e az elsõ bemenet. Ha nem akkor hiba van.
    If Not IsNumeric(bemenet1.Text) Then
        vegeredmeny.Alignment = 0                                 ' Szöveg balra zárása.
        vegeredmeny.Text = "Hiba! Az elsõ bemenetbe nem szám lett megadva!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilpés a hibakeresésbõl.
    End If

    ' Megnézzük szám-e a második bemenet. Ha nem akkor hiba van.
    ' Továbbá ha négyzetre emelés vagy gyökvonás van akkor nem fut le ez a feltétel.
    If Not IsNumeric(bemenet2.Text) And Not nincsbemenet2 Then
        vegeredmeny.Alignment = 0                                 ' Szöveg balra zárása.
        vegeredmeny.Text = "Hiba! A második bemenetbe nem szám lett megadva!"
        Hibakereses = True                                        ' Hiba van.
        Exit Function                                             ' Kilpés a hibakeresésbõl.
    End If

    Hibakereses = False                                           ' Nincs hiba.
End Function
