VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* SelfTimer 1.0 - Timer class module
'* ----------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) Free for commercial and non-commercial usage.
'*        4) Use at your own risk. No support guaranteed.
'*
'* REQUIREMENTS
'* ------------
'* Huh what? Just this one class module. No extra files required.
'*
'* HOW TO ADD TO YOUR PROGRAM
'* --------------------------
'* 1) Copy SelfTimer.cls to your project folder.
'* 2) In your project, add SelfTimer.cls
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.0 (2008-06-15)
'* - It was requested to simplify SelfCallback code and make a one class module timer at VBForums.
'*   I renamed and cleaned up the code to fit my own coding style. Good or bad, a matter of taste.
'*
'* CREDITS
'* -------
'* Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'*************************************************************************************************
Option Explicit

' events
Public Event Timer(ByVal Seconds As Currency)

' public properties
Private m_Enabled As Boolean
Private m_Interval As Long

' private variables
Private m_Start As Currency
Private m_TimerID As Long
Private m_TimerProc As Long

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

' the following are the requirements for SelfCallback
Private Const IDX_CALLBACKORDINAL As Long = 36

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    m_Enabled = NewValue
    ' we can call these both; the order is important, of course
    TimerKill
    TimerSet
End Property
Public Property Get Interval() As Long
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal NewValue As Long)
    ' reset starting time (we do this every time Interval changes)
    m_Start = 0
    ' the value can't be negative
    If NewValue > 0 Then
        ' kill existing timer?
        TimerKill
        ' update value
        m_Interval = NewValue
        ' start timer
        TimerSet
    ElseIf NewValue = 0 Then
        ' kill existing timer?
        TimerKill
        ' update value
        m_Interval = 0
    End If
End Property
Private Function ScAddressOfOrdinal(ByVal Ordinal As Long) As Long
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long, lngAddress As Long
    ' get address of this class module instance
    RtlMoveMemory VarPtr(lngAddress), ObjPtr(Me), 4&
    ' probe for Class method
    If ScProbe(lngAddress + &H1C, lngA, bytSignature) Then
        ' scan up to 256 vTable entries
        For lngA = lngA + 4 To lngA + 1024 Step 4
            ' get address in vTable
            RtlMoveMemory VarPtr(lngAddress), lngA, 4&
            ' invalid code address?
            If IsBadCodePtr(lngAddress) Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
            ' get byte pointed to vTable address
            RtlMoveMemory VarPtr(bytValue), lngAddress, 1&
            ' if does not match the expected value...
            If bytValue <> bytSignature Then
                ' return this vTable address
                RtlMoveMemory VarPtr(ScAddressOfOrdinal), lngA - (Ordinal * 4&), 4&
                Exit For
            End If
        Next lngA
    End If
End Function
Private Property Get ScData(ByVal Index As Long, ByVal ThunkPtr As Long) As Long
    RtlMoveMemory VarPtr(ScData), ThunkPtr + (Index * 4&), 4&
End Property
Private Property Let ScData(ByVal Index As Long, ByVal ThunkPtr As Long, ByVal NewValue As Long)
    RtlMoveMemory ThunkPtr + (Index * 4&), VarPtr(NewValue), 4&
End Property
Private Function ScProbe(ByVal Address As Long, ByRef Method As Long, ByRef Signature As Byte) As Boolean
    Dim bytValue As Byte, lngVTableEntry As Long
    ' probe eight entries
    For Address = Address To Address + 32 Step 4
        ' get vTable entry
        RtlMoveMemory VarPtr(lngVTableEntry), Address, 4&
        ' if not an implemented interface
        If lngVTableEntry Then
            ' get the value pointed at by the vTable entry
            RtlMoveMemory VarPtr(bytValue), lngVTableEntry, 1&
            ' if native or P-code signature...
            If (bytValue = &H33) Or (bytValue = &HE9) Then
                ' return this information
                Method = Address
                Signature = bytValue
                ' success, exit loop
                ScProbe = True
                Exit For
            End If
        End If
    Next Address
End Function
Private Function ScProcedureAddress(ByVal DynamicLinkLibrary As String, ByVal Procedure As String, ByVal Unicode As Boolean) As Long
    ' get the procedure address
    If Unicode Then
        ScProcedureAddress = GetProcAddress(GetModuleHandleW(StrPtr(DynamicLinkLibrary)), Procedure)
    Else
        ScProcedureAddress = GetProcAddress(GetModuleHandleA(DynamicLinkLibrary), Procedure)
    End If
    ' in IDE, verify we got it
    Debug.Assert ScProcedureAddress
End Function
Private Sub TimerKill()
    Const MEM_RELEASE As Long = &H8000&
    ' only do this if we still have the procedure
    If m_TimerProc Then
        ' kill the timer
        KillTimer 0&, m_TimerID
        ' reset id
        m_TimerID = 0
        ' free the procedure callback
        VirtualFree m_TimerProc, 0, MEM_RELEASE
        ' reset procedure pointer to prevent this getting ran twice
        m_TimerProc = 0
    End If
End Sub
Private Sub TimerSet()
    Dim lngCallback As Long, lngCallbackCode() As Long, lngScMem As Long

    ' index of ordinal (always 1 in this class, the last procedure)
    Const Ordinal As Long = 1
    ' number of parameters (for TimerProc this is always 4, thus this is a constant)
    Const ParamCount As Long = 4

    ' memory bytes required for the callback thunk
    Const MEM_LEN As Long = IDX_CALLBACKORDINAL * 4 + 4
    ' allocate executable memory
    Const PAGE_RWX As Long = &H40&
    ' commit allocated memory
    Const MEM_COMMIT As Long = &H1000&
    ' name of this routine
    Const SUB_NAME As String = "TimerSet"
    ' thunk data index of the Owner object's vTable address
    Const INDX_OWNER As Long = 0
    ' thunk data index of the callback procedure address
    Const INDX_CALLBACK As Long = 1
    ' thunk data index of the EbMode function address
    Const INDX_EBMODE As Long = 2
    ' thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR As Long = 3
    ' thunk data index of the KillTimer function address
    Const INDX_KT As Long = 4
    ' thunk code patch index of the thunk data
    Const INDX_EBX As Long = 6
    ' thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMS As Long = 18
    ' thunk code patch index of the bytes to be released after callback
    Const INDX_PARAMLEN As Long = 24
    ' thunk offset to the callback execution address
    Const PROC_OFF As Long = &H14

    ' start only if enabled, interval is set and we have not done this already
    If m_TimerProc = 0 And m_Enabled And m_Interval <> 0 Then
        ' get address of the TimerProc procedure, the last procedure in this class module (by ordinal)
        lngCallback = ScAddressOfOrdinal(Ordinal)
        ' verify we got it
        If lngCallback Then
            ' allocate executable memory
            lngScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
            ' verify we got it
            If lngScMem Then
                ' allocate for the machine-code array
                ReDim lngCallbackCode(0 To IDX_CALLBACKORDINAL) As Long
                ' create the machine-code array
                lngCallbackCode(5) = &HBB60E089
                lngCallbackCode(7) = &H73FFC589
                lngCallbackCode(8) = &HC53FF04
                lngCallbackCode(9) = &H59E80A74
                lngCallbackCode(10) = &HE9000000
                lngCallbackCode(11) = &H30&
                lngCallbackCode(12) = &H87B81
                lngCallbackCode(13) = &H75000000
                lngCallbackCode(14) = &H9090902B
                lngCallbackCode(15) = &H42DE889
                lngCallbackCode(16) = &H50000000
                lngCallbackCode(17) = &HB9909090
                lngCallbackCode(19) = &H90900AE3
                lngCallbackCode(20) = &H8D74FF
                lngCallbackCode(21) = &H9090FAE2
                lngCallbackCode(22) = &H53FF33FF
                lngCallbackCode(23) = &H90909004
                lngCallbackCode(24) = &H2BADC261
                lngCallbackCode(25) = &H3D0853FF
                lngCallbackCode(26) = &H1&
                lngCallbackCode(27) = &H23DCE74
                lngCallbackCode(28) = &H74000000
                lngCallbackCode(29) = &HAE807
                lngCallbackCode(30) = &H90900000
                lngCallbackCode(31) = &H4589C031
                lngCallbackCode(32) = &H90DDEBFC
                lngCallbackCode(33) = &HFF0C75FF
                lngCallbackCode(34) = &H53FF0475
                lngCallbackCode(35) = &HC310&
                ' settings
                lngCallbackCode(INDX_BADPTR) = ScProcedureAddress("kernel32", "IsBadCodePtr", False)
                ' set owner (this class)
                lngCallbackCode(INDX_OWNER) = ObjPtr(Me)
                ' callback address (the last procedure in this class)
                lngCallbackCode(INDX_CALLBACK) = lngCallback
                ' remember the ordinal used
                lngCallbackCode(IDX_CALLBACKORDINAL) = Ordinal
                ' parameter count
                lngCallbackCode(INDX_PARAMS) = ParamCount
                RtlMoveMemory VarPtr(lngCallbackCode(INDX_PARAMLEN)) + 2&, VarPtr(ParamCount * 4&), 2&
                ' IDE safety
                If App.LogMode = 0 Then
                    ' store the EbMode function address in the thunk data
                    lngCallbackCode(INDX_EBMODE) = ScProcedureAddress("vba6", "EbMode", False)
                End If
                ' special for timer callback:
                lngCallbackCode(INDX_KT) = ScProcedureAddress("user32", "KillTimer", False)
                ' set the data address relative to virtual memory pointer
                lngCallbackCode(INDX_EBX) = lngScMem
                ' copy thunk code to executable memory
                RtlMoveMemory lngScMem, VarPtr(lngCallbackCode(INDX_OWNER)), MEM_LEN
                ' remember the procedure address
                m_TimerProc = lngScMem + PROC_OFF
                ' now we can initialize the timer
                m_TimerID = SetTimer(0&, 0&, m_Interval, m_TimerProc)
                ' done!
                Exit Sub
            End If
        End If
        ' timer initialization failed for whatever reason, thus timer is disabled
        m_Enabled = False
    End If
End Sub
Private Sub Class_Initialize()
    ' initial values: we have it enabled but we have no interval
    m_Enabled = True
    m_Interval = 0
End Sub
Private Sub Class_Terminate()
    TimerKill
End Sub
' must be the last procedure, ordinal #1!
Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    Dim curCounter As Currency
    ' copy Long to Currency
    RtlMoveMemory VarPtr(curCounter), VarPtr(tickCount), 4&
    ' see if start has been initialized (it never will be zero)
    If m_Start > 0 Then
        ' calculate seconds since beginning
        RaiseEvent Timer(curCounter * 10 - m_Start)
    Else
        ' just remember the start time and start from zero
        m_Start = curCounter * 10
        RaiseEvent Timer(0)
    End If
    ' WHATEVER YOU DO, DO NOT CALL TimerKill FROM HERE!
End Function

