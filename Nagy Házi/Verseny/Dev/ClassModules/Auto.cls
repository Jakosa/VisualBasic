VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Timer0 As VB.Timer
Attribute Timer0.VB_VarHelpID = -1
Private WithEvents KVonal1 As VB.Line
Attribute KVonal1.VB_VarHelpID = -1
Private WithEvents KVonal2 As VB.Line
Attribute KVonal2.VB_VarHelpID = -1
Private WithEvents KVonal3 As VB.Line
Attribute KVonal3.VB_VarHelpID = -1
Private WithEvents KVonal4 As VB.Line
Attribute KVonal4.VB_VarHelpID = -1
Private VSzogTomb(1 To 30) As Double
Private Color As ColorConstants
Private Nyomvonal As Boolean
Private Started As Boolean
Private x0 As Double
Private y0 As Double
Private ex As Double
Private ey As Double
Private i As Long
' Ha "true" akkor tökéletesebben középen van az autó. "false" esetén viszont szerintem élethübb.
Private PontosabbKor As Boolean
' Kocsi hossza.
Private Const xhossz = 80
' Kocsi szélessége.
Private Const yhossz = 60
Private KorokSzama As Byte
' Egy kör idejének hossza.
Private EgyKorido As Date
Private StopKorido As Date
' Induláskor a köridõ beállításában segítõ változó.
Private BKorido As Boolean
Private BStopKorido As Boolean
' Egy szektor idejének hossza.
Private EgySzektorido As Date
' Tárolja hogy az autó átment-e az egyik szetoron vagy sem.
Private ASzektorvonal As Boolean
Private LegjobbSzektorido(1 To 3) As Single
Private LegjobbKorido As Single
Private LegjobbKoridoSzama As Byte
Private AutoSzama As Byte
' Tárolja vége van-e a játéknak vagy sem.
Private GameEnd As Boolean
Private UtolsoMeterek As Integer
Private EgyKorHossza As Integer
Private OsszesUt As Integer
' Segití a startvonal elõli indulást.
Private IndulasStartVonalElol As Boolean

Private Type Korok
    Korido As Single
    Szektorido(1 To 3) As Single
End Type

Private KorokTomb() As Korok

Public Property Get GetGameEnd() As Boolean
    GetGameEnd = GameEnd
End Property

Public Property Get GetEgyKorHossza() As Integer
    GetEgyKorHossza = EgyKorHossza
End Property

Public Property Get GetOsszesUt() As Integer
    GetOsszesUt = OsszesUt
End Property

Public Property Get GetKorokSzama() As Byte
    GetKorokSzama = KorokSzama
End Property

Public Property Get GetKorSzektoridok(ii As Integer) As Single
    GetKorSzektoridok = KorokTomb(KorokSzama).Szektorido(ii)
End Property

Public Property Get GetOsszKorido() As Single
    Dim osszido As Single, ciklus As Integer

    For ciklus = Palya.GetKezdokorErteke To Config.Globalis_KorokSzama
        osszido = osszido + KorokTomb(ciklus).Korido
    Next ciklus

    GetOsszKorido = osszido
End Property

Public Property Get GetLegjobbKorido() As Single
    GetLegjobbKorido = LegjobbKorido
End Property

Public Property Get GetLegjobbKoridoSzama() As Byte
    GetLegjobbKoridoSzama = LegjobbKoridoSzama
End Property

Public Property Get GetLegjobbSzektoridok(ii As Integer) As Single
    GetLegjobbSzektoridok = LegjobbSzektorido(ii)
End Property

Public Property Get GetKVonal1() As VB.Line
    Set GetKVonal1 = KVonal1
End Property

Public Property Get GetKVonal2() As VB.Line
    Set GetKVonal2 = KVonal2
End Property

Public Property Get GetKVonal3() As VB.Line
    Set GetKVonal3 = KVonal3
End Property

Public Property Get GetKVonal4() As VB.Line
    Set GetKVonal4 = KVonal4
End Property

Public Property Get GetColor() As String
    Dim Szin As String

    Select Case Color
        Case ColorConstants.vbBlue
            Szin = "kék"
        Case ColorConstants.vbRed
            Szin = "piros"
        Case ColorConstants.vbGreen
            Szin = "zöld"
        Case ColorConstants.vbYellow
            Szin = "sárga"
        Case ColorConstants.vbBlack
            Szin = "fekete"
        Case Else
            Szin = "fekete"
    End Select

    GetColor = Szin
End Property

Public Property Get GetX0() As Double
    GetX0 = x0
End Property

Public Property Get GetY0() As Double
    GetY0 = y0
End Property

Public Sub SetX0(ByVal X As Double)
    x0 = KorrigacioX(X)
End Sub

Public Sub SetY0(ByVal Y As Double)
    y0 = KorrigacioY(Y)
End Sub

Public Sub SetEX(ByVal X As Double)
    ex = X
End Sub

Public Sub SetEY(ByVal Y As Double)
    ey = Y
End Sub

Public Sub SetColor(ByVal Szin As String)
    Select Case Szin
        Case "Kék", "kék"
            Color = ColorConstants.vbBlue
        Case "Piros", "piros"
            Color = ColorConstants.vbRed
        Case "Zöld", "zöld"
            Color = ColorConstants.vbGreen
        Case "Sárga", "sárga"
            Color = ColorConstants.vbYellow
        Case "Fekete", "fekete"
            Color = ColorConstants.vbBlack
        Case Else
            Color = ColorConstants.vbBlack
    End Select

    KVonal1.BorderColor = Color
    KVonal2.BorderColor = Color
    KVonal3.BorderColor = Color
    KVonal4.BorderColor = Color
End Sub

Public Sub SetVisible(ByVal Visible As Boolean)
    KVonal1.Visible = Visible
    KVonal2.Visible = Visible
    KVonal3.Visible = Visible
    KVonal4.Visible = Visible
End Sub

Public Sub SetBorderWidth(ByVal Nagysag As Byte)
    KVonal1.BorderWidth = Nagysag
    KVonal2.BorderWidth = Nagysag
    KVonal3.BorderWidth = Nagysag
    KVonal4.BorderWidth = Nagysag
End Sub

Public Sub SetNyomvonal(ByVal Nyom As Boolean)
    Nyomvonal = Nyom
End Sub

Public Sub SetPontosabbKor(ByVal PKor As Boolean)
    PontosabbKor = PKor
End Sub

Public Sub Load(ByVal Szam As Byte)
    AutoSzama = Szam
    Set KVonal1 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "1", Palya.VirtualisPalya)
    Set KVonal2 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "2", Palya.VirtualisPalya)
    Set KVonal3 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "3", Palya.VirtualisPalya)
    Set KVonal4 = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & "4", Palya.VirtualisPalya)

    SetVisible True

    Set Timer0 = Palya.Controls.Add("VB.Timer", "Timer" & Szam & "1", Palya)
    Timer0.Enabled = False
    Timer0.Interval = 50 + Rnd * 21

    Dim ciklus As Long
    For ciklus = 1 To 30
        VSzogTomb(ciklus) = Rnd * 0.1
    Next ciklus
End Sub

Public Sub Start()
    If Started Or GameEnd Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    If BKorido Then
        EgyKorido = Now
        EgySzektorido = EgyKorido
        BKorido = False
    End If

    Started = True
    Timer0.Enabled = True
End Sub

Public Sub Stop_Kocsi()
    If Not Started Or GameEnd Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    Started = False
    StopKorido = Now
    BStopKorido = True
    Timer0.Enabled = False
End Sub

Private Sub Class_Initialize()
    Dim ciklus As Integer

    ex = 0.6
    ey = -1
    x0 = KorrigacioX(1100)
    y0 = KorrigacioY(5000)
    BKorido = True
    GameEnd = False
    BStopKorido = False
    IndulasStartVonalElol = False
    KorokSzama = Palya.GetKezdokorErteke()
    SetNyomvonal True
    SetPontosabbKor False
    ASzektorvonal = False
    ReDim KorokTomb(Palya.GetKezdokorErteke() To Config.Globalis_KorokSzama) As Korok

    LegjobbKorido = KezdoSzektorido   ' Kezdõérték

    For ciklus = LBound(LegjobbSzektorido) To UBound(LegjobbSzektorido)
        LegjobbSzektorido(ciklus) = KezdoSzektorido  ' Kezdõérték
    Next ciklus
End Sub

Private Sub Class_Terminate()
    ' Nullázás.
    Set KVonal1 = Nothing
    ' Nullázás.
    Set KVonal2 = Nothing
    ' Nullázás.
    Set KVonal3 = Nothing
    ' Nullázás.
    Set KVonal4 = Nothing
    ' Nullázás.
    Set Timer0 = Nothing
End Sub

Public Sub Dispose()
    Dim i As Long
    For i = 1 To 4
        Palya.Controls.Remove "KocsiVonal" & AutoSzama & i
    Next i

    Palya.Controls.Remove "Timer" & AutoSzama & "1"
End Sub

Public Sub Show()
    NextMove
End Sub

' Korrigálja az X eltérést.
Private Function KorrigacioX(ByVal PointX As Single) As Single
    ' A virtuális pálya x tengely távolságának levonása.
    KorrigacioX = PointX - Palya.VirtualisPalya.Left
End Function

' Korrigálja az Y eltérést.
Private Function KorrigacioY(ByVal PointY As Single) As Single
    ' A virtuális pálya y tengely távolságának levonása.
    KorrigacioY = PointY - Palya.VirtualisPalya.Top
End Function

' Két szektoridõ mérése értelemszerüen távolság alapján ha elérte.
' 0-tól 1-ig tartanak a szektorvonalak mert a célvonal egyben a harmadik szektor.
Private Sub Szektorok()
    If KorokSzama > Config.Globalis_KorokSzama Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    Dim NextKor As Boolean, ii As Byte
    For ii = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido)
        Dim dist As Single, ciklus As Byte, Szam As Byte
        dist = 1000000

        For ciklus = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
            Dim d As Single
            d = Distance(x0, y0, KorrigacioX(PalyaInfo.SzektorVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.SzektorVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.SzektorVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.SzektorVonalTomb(ciklus).Y2))

            If dist > d Then
                dist = d
                Szam = ciklus + 1 ' Mert egyel magasabbon vannak
            End If
        Next ciklus

        If dist < 100 And Not ASzektorvonal Then
            Dim NowTime As Date
            NowTime = Now
            ASzektorvonal = True

            If Not IndulasStartVonalElol Then
                IndulasStartVonalElol = True
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            For ciklus = LBound(PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok) To UBound(PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok)
                If Not PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).VanAdat Then
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).VanAdat = True
                End If

                If PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Szin = "" Then
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Szin = GetColor
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Ido = NowTime
                    ' Kilépés a ciklusból.
                    Exit For
                End If
            Next ciklus

            If Szam = 3 Then
                NextKor = True
                KorokTomb(KorokSzama).Korido = DateDiff("s", EgyKorido, NowTime)

                If LegjobbKorido > KorokTomb(KorokSzama).Korido Then
                    LegjobbKorido = KorokTomb(KorokSzama).Korido
                    LegjobbKoridoSzama = KorokSzama
                End If

                KorokTomb(KorokSzama).Szektorido(Szam) = KorokTomb(KorokSzama).Korido

                Dim j As Byte
                For j = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido) - 1 ' Mert utolsót nem akarjuk kivonni
                    KorokTomb(KorokSzama).Szektorido(Szam) = KorokTomb(KorokSzama).Szektorido(Szam) - KorokTomb(KorokSzama).Szektorido(j)
                Next j
            Else
                KorokTomb(KorokSzama).Szektorido(Szam) = DateDiff("s", EgySzektorido, NowTime)
            End If

            If LegjobbSzektorido(Szam) > KorokTomb(KorokSzama).Szektorido(Szam) Then
                LegjobbSzektorido(Szam) = KorokTomb(KorokSzama).Szektorido(Szam)
            End If

            NowTime = Now

            If Szam = 3 Then
                EgyKorido = NowTime
            End If

            EgySzektorido = NowTime
        ElseIf dist > 100 And ASzektorvonal Then
            ASzektorvonal = False
        End If
    Next ii

    If NextKor Then
        KorokSzama = KorokSzama + 1
    End If

    If KorokSzama > Config.Globalis_KorokSzama Then
        GameEnd = True
    End If
End Sub

Private Sub Timer0_Timer()
    i = i + 1

    If i > 30 Then
        i = 1

        Randomize
        Timer0.Interval = 50 + Rnd * 21
    End If

    If BStopKorido Then
        Dim NowTime As Date
        NowTime = Now
        EgyKorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgyKorido)
        EgySzektorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgySzektorido)
        BStopKorido = False
    End If

    NextCoordinate True, True

    Dim xb As Single, yb As Single, xj As Single, yj As Single
    ' Ideálisan középen van így a jármû.
    If PontosabbKor Then
        xb = x0 + xhossz * ex - yhossz * ey
        yb = y0 + xhossz * ey + yhossz * ex
        xj = x0 + xhossz * ex + yhossz * ey
        yj = y0 + xhossz * ey - yhossz * ex
    Else
        xb = x0 + 400 * ex - 300 * ey
        yb = y0 + 400 * ey + 300 * ex
        xj = x0 + 400 * ex + 300 * ey
        yj = y0 + 400 * ey - 300 * ex
    End If

    Dim distb As Single, distj As Single, ciklus As Integer
    distb = 1000000
    distj = 1000000

    For ciklus = LBound(PalyaInfo.PalyaVonalTomb) To PalyaInfo.PalyaVonalakSzama - 1
        Dim db As Single, dj As Single
        db = Distance(xb, yb, KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y2))
        dj = Distance(xj, yj, KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y2))

        If distb > db Then
            distb = db
        End If

        If distj > dj Then
            distj = dj
        End If
    Next ciklus

    ' Ideálisan középen van így a jármû.
    If PontosabbKor Then
        If distb - distj > 0 And distb - distj < 200 Then
            'If distb - distj > 0 And distb - distj < 50 Then
            '    NextCoordinate False, False
            '    NextMove
            'End If

            NextCoordinate False, False
            NextMove
        ElseIf distj - distb > 0 And distj - distb < 200 Then
            'If distj - distb > 0 And distj - distb < 50 Then
            '    NextCoordinate True, False
            '    NextMove
            'End If

            NextCoordinate True, False
            NextMove
        End If
    Else
        If distb < 200 Then
            NextCoordinate True, False
            NextMove
        ElseIf distb > 200 Then
            NextCoordinate False, False
            NextMove
        Else
            x0 = x0 - 55 * ex
            y0 = y0 - 55 * ey
        End If
    End If

    NextMove

    x0 = x0 + 50 * ex
    y0 = y0 + 50 * ey
    NextMove

    Szektorok

    If Nyomvonal Then
        Palya.VirtualisPalya.Circle (x0, y0), 10, Color
    End If

    If KorokSzama = Palya.GetKezdokorErteke And IndulasStartVonalElol Then
        EgyKorHossza = EgyKorHossza + PalyaHosszanakLepteke ' Méterben van értve.
    End If

    OsszesUt = OsszesUt + PalyaHosszanakLepteke         ' Méterben van értve.

    If GameEnd Then
        UtolsoMeterek = UtolsoMeterek + 1

        If UtolsoMeterek = 20 Then
            Timer0.Enabled = False
        End If
    End If
End Sub

Private Sub NextCoordinate(ByVal BalraKanyarodas As Boolean, ByVal VSzog As Boolean)
    Dim bj As Double

    If BalraKanyarodas Then
        bj = -0.15
    Else
        bj = 0.15
    End If

    If VSzog Then
        ex = Cos(VSzogTomb(i)) * ex - Sin(VSzogTomb(i)) * ey
        ey = Cos(VSzogTomb(i)) * ey + Sin(VSzogTomb(i)) * ex
    Else
        ex = Cos(bj) * ex - Sin(bj) * ey
        ey = Cos(bj) * ey + Sin(bj) * ex
    End If

    ex = ex / Sqr(ex * ex + ey * ey)
    ey = ey / Sqr(ex * ex + ey * ey)
End Sub

' Mozgatás a következõ pontokba.
Private Sub NextMove()
    ' Vonal elsõ X pontjának legenerálása.
    KVonal1.X1 = x0 + xhossz * ex - yhossz * ey
    ' Vonal elsõ Y pontjának legenerálása.
    KVonal1.Y1 = y0 + xhossz * ey + yhossz * ex
    ' Vonal második X pontjának legenerálása.
    KVonal1.X2 = x0 - xhossz * ex - yhossz * ey
    ' Vonal második Y pontjának legenerálása.
    KVonal1.Y2 = y0 - xhossz * ey + yhossz * ex

    ' Vonal elsõ X pontjának legenerálása.
    KVonal2.X1 = x0 + xhossz * ex + yhossz * ey
    ' Vonal elsõ Y pontjának legenerálása.
    KVonal2.Y1 = y0 + xhossz * ey - yhossz * ex
    ' Vonal második X pontjának legenerálása.
    KVonal2.X2 = x0 - xhossz * ex + yhossz * ey
    ' Vonal második Y pontjának legenerálása.
    KVonal2.Y2 = y0 - xhossz * ey - yhossz * ex

    ' Vonal elsõ X pontjának átadása.
    KVonal3.X1 = KVonal1.X1
    ' Vonal második X pontjának átadása.
    KVonal3.X2 = KVonal2.X1
    ' Vonal elsõ Y pontjának átadása.
    KVonal3.Y1 = KVonal1.Y1
    ' Vonal második Y pontjának átadása.
    KVonal3.Y2 = KVonal2.Y1

    ' Vonal elsõ X pontjának átadása.
    KVonal4.X1 = KVonal1.X2
    ' Vonal második X pontjának átadása.
    KVonal4.X2 = KVonal2.X2
    ' Vonal elsõ Y pontjának átadása.
    KVonal4.Y1 = KVonal1.Y2
    ' Vonal második Y pontjának átadása.
    KVonal4.Y2 = KVonal2.Y2
End Sub
