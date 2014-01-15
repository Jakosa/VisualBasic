Attribute VB_Name = "Calc"
Option Explicit

Public Const MAXTOKENS = 1000

Public Const TKN_UNKNOWN = "TKN_UNKNOWN"
Public Const TKN_NUMBER = "TKN_NUMBER"
Public Const TKN_OPERATOR = "TKN_OPERATOR"
Public Const TKN_NAME = "TKN_NAME"
Public Const TKN_OPEN_PAR = "TKN_OPEN_PAR"
Public Const TKN_CLOSE_PAR = "TKN_CLOSE_PAR"
Public Const TKN_TERMINATION = "TKN_TERMINATION"
Public Const TKN_NEXT_LINE = "TKN_NEXT_LINE"

Public Function getpriority(tk As Token) As Integer
    Debug.Print "tk.str: " & tk.str

    If tk.str = "^" Then
        getpriority = 5
    ElseIf tk.str = "**" Then
        getpriority = 5
    ElseIf tk.str = "*" Then
        getpriority = 6
    ElseIf tk.str = "/" Then
        getpriority = 6
    ElseIf tk.str = "%" Then
        getpriority = 6
    ElseIf tk.str = "+" Then
        getpriority = 7
    ElseIf tk.str = "-" Then
        getpriority = 7
    'ElseIf tk.str = ":" Then
    '    getpriority = 16
    ElseIf tk.str = "(" Then
        getpriority = 100
    Else
        getpriority = 0
    End If
End Function

' Inéttõl olyan függvények vannak amiket netrõl vettem mert nem akartam szenvedni a kifejlesztésével.
Public Sub spush(ArrayName() As String, Element As String)
    ' /* Make Sure The Variable Passed Is An Array */
    If IsArray(ArrayName) = False Then Exit Sub
    ' /* If We Get An Error Here, it's Probably Because the Array */
    ' /* is dimensioned with no values set, so just make the first element */
    On Error GoTo make_it
    ' /* Allocate A New Array Slot */
    ReDim Preserve ArrayName(UBound(ArrayName()) + 1)
    ' /* Set The Value Of The New Array Indice To Element */
    ArrayName(UBound(ArrayName())) = Element
    Exit Sub
    ' /* Code Will Only Jump Here "on error" */
make_it:
    ReDim ArrayName(0)
    ArrayName(0) = Element
End Sub

Public Function str2Array(xString As String) As String()
    Dim tmpArray() As String
    Dim tmpchar As String
    Dim i As Long
    
    ' /* For Each Character In The String */
    For i = 1 To Len(xString)
        ' /* Retrieve The Character */
        tmpchar = Mid(xString, i, 1)
        ' /* Push It Into The Temporary Array */
        spush tmpArray, tmpchar
    Next i
    
    ' /* Return The Array To The Calling Procedure */
    str2Array = tmpArray
End Function
