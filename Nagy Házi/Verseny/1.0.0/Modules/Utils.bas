Attribute VB_Name = "Utils"
' Fejléc
' Készítette: Jakosa Csaba Árpád
' Fejléc vége

Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As EnSetWindowPos) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function SetViewportOrgEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Const WM_PAINT As Long = &HF&

' Létrehozza a Config változót.
Public Config As New Config
' Létrehozza a Map változót.
Public Map As New Map
' Létrehozza a VegeredmenyMentese változót.
Public VegeredmenyMentese As New VegeredmenyMentese

' Megnézi hogy a változó "boolean"-e.
' A "value" tárolja változó értékét.
Public Function IsBoolean(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibarészhez.
    On Error GoTo ErrorHandler

    ' b változó létrehozása.
    Dim b As Boolean
    ' Megnézi hogy a "value" értéke igaz-e vagy hamis. Valóságban megpróbálja konvertálni.
    b = CBool(value)
    ' Ha sikeres a konvertálás akkor nincs hiba ezért igaz értékre állítja a visszatérési értéket.
    IsBoolean = True

    ' Hiba esetén ide ugrik a program.
ErrorHandler:
    ' Hamis (False) értékkel lép ki.
End Function

' Megnézi hogy a változó "integer"-e.
' A "value" tárolja változó értékét.
Public Function IsInteger(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibarészhez.
    On Error GoTo ErrorHandler

    ' i változó létrehozása.
    Dim i As Boolean
    ' Megnézi hogy a "value" értéke igaz-e vagy hamis. Valóságban megpróbálja konvertálni.
    i = CInt(value)
    ' Ha sikeres a konvertálás akkor nincs hiba ezért igaz értékre állítja a visszatérési értéket.
    IsInteger = True

    ' Hiba esetén ide ugrik a program.
ErrorHandler:
    ' Hamis (False) értékkel lép ki.
End Function

' Megnézi hogy a változó "byte"-e.
' A "value" tárolja változó értékét.
Public Function IsByte(ByVal value As Variant) As Boolean
    ' Ha hiba van akkor ugrik a hibarészhez.
    On Error GoTo ErrorHandler

    ' b változó létrehozása.
    Dim b As Boolean
    ' Megnézi hogy a "value" értéke igaz-e vagy hamis. Valóságban megpróbálja konvertálni.
    b = CByte(value)
    ' Ha sikeres a konvertálás akkor nincs hiba ezért igaz értékre állítja a visszatérési értéket.
    IsByte = True

    ' Hiba esetén ide ugrik a program.
ErrorHandler:
    ' Hamis (False) értékkel lép ki.
End Function

' Kiszámolja a pont és szakasz távolságát.
' A "PointX" tárolja a pont X koordinátáját.
' A "PointY" tárolja a pont Y koordinátáját.
' A "LineX1" tárolja a szakasz X1 koordinátáját.
' A "LineX2" tárolja a szakasz X2 koordinátáját.
' A "LineY1" tárolja a szakasz Y1 koordinátáját.
' A "LineY2" tárolja a szakasz Y2 koordinátáját.
Public Function Distance(ByVal PointX As Single, ByVal PointY As Single, ByVal LineX1 As Single, ByVal LineX2 As Single, ByVal LineY1 As Single, ByVal LineY2 As Single) As Single
    ' Változók.
    Dim AA As Single, BB As Single, CC As Single, DD As Single
    ' Változók.
    Dim dot As Single, len_sq As Single, param As Single
    ' Változók.
    Dim xx As Single, yy As Single

    ' X pont és egyenes X1 koordinátájának különbsége.
    AA = PointX - LineX1
    ' Y pont és egyenes Y1 koordinátájának különbsége.
    BB = PointY - LineY1
    ' Egyenes X2 és egyenes X1 koordinátájának különbsége.
    CC = LineX2 - LineX1
    ' Egyenes Y2 és egyenes Y1 koordinátájának különbsége.
    DD = LineY2 - LineY1

    ' Változók szorzata és összege.
    dot = AA * CC + BB * DD
    ' Változók szorzata és összege.
    len_sq = CC * CC + DD * DD
    ' Két változó hányadosa.
    param = dot / len_sq

    ' Akkor fut le ha a param kisebb mint 0.
    If param < 0 Then
        ' xx beállítása.
        xx = LineX1
        ' yy beállítása.
        yy = LineY1
    ' Akkor fut le ha a param nagyobb mint 1.
    ElseIf param > 1 Then
        ' xx beállítása.
        xx = LineX2
        ' yy beállítása.
        yy = LineY2
    Else
        ' xx beállítása.
        xx = LineX1 + param * CC
        ' yy beállítása.
        yy = LineY1 + param * DD
    End If

    ' Visszadja a pontok távolságát.
    Distance = Sqr(((PointX - xx) * (PointX - xx)) + ((PointY - yy) * (PointY - yy)))
End Function

' Megnyítja és beállítja a hibaablak paramétereit.
' A "Title" az ablak címe.
' A "Message" az ablakra kiírt hibaüzenet.
' A "Leallas" pedig a program leállítását tárolja. Ha szükséges.
Public Sub WarningWindow(ByVal Title As String, ByVal Message As String, ByVal Leallas As Boolean)
    ' Hibaüzenet eltárolása.
    WarningForm.HibaUzenet = Message
    ' Leállítás eltárolása.
    WarningForm.Leallitas = Leallas
    ' Ablak címmének eltárolása.
    WarningForm.Caption = Title
    ' Megjeleniti a hiba ablakot. Közben meggátolja hogy a mögötte lévõ ablakra rá lehessen kattintani (1 a tiltás).
    WarningForm.Show 1
End Sub
