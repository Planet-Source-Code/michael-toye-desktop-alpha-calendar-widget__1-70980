VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19290
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   19290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private CurrMonth&, CurrYear&, CurrAlpha&

Sub SetAlpha(Opac&)
    Dim Ret&
    Const LWA_COLORKEY = &H1
    Const LWA_ALPHA = &H2
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, 0, Opac, LWA_ALPHA
End Sub



Private Sub Form_DblClick()
SetWindowTopmost Me.hwnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 65
            CurrAlpha = CurrAlpha + 50
            If CurrAlpha > 255 Then CurrAlpha = CurrAlpha - 255
            SetAlpha CurrAlpha
            
        Case 39
            CurrMonth = CurrMonth + 1
            If CurrMonth > 12 Then
                CurrMonth = 1
                CurrYear = CurrYear + 1
            End If
            DrawScreen
            
        Case 37
            CurrMonth = CurrMonth - 1
            If CurrMonth < 1 Then
                CurrMonth = 12
                CurrYear = CurrYear - 1
            End If
            DrawScreen
            
        Case 40
            CurrYear = CurrYear - 1
            DrawScreen
            
        Case 38
            CurrYear = CurrYear + 1
            DrawScreen
            
    End Select
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton And X > 210 And X < (Me.Width - 210) And Y < 360 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
ElseIf X < 210 Then
    CurrMonth = CurrMonth - 1
    If CurrMonth <= 0 Then
        CurrMonth = 12
        CurrYear = CurrYear - 1
    End If
    DrawScreen
ElseIf X > (Me.Width - 210) Then
    CurrMonth = CurrMonth + 1
    If CurrMonth > 12 Then
        CurrMonth = 1
        CurrYear = CurrYear + 1
    End If
    DrawScreen
ElseIf Y > (Me.Height - 100) Then
    Unload Me
    End
End If
End Sub
Private Sub SetWindowTopmost(hwnd&)
Static OnTop As Boolean

    If OnTop Then
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    OnTop = Not OnTop
    
End Sub
Private Sub SplitRGB(ByVal clr&, R&, G&, B&)
    R = clr And &HFF: G = (clr \ &H100&) And &HFF: B = (clr \ &H10000) And &HFF
End Sub
Private Sub Gradient(DC&, X&, Y&, dx&, dy&, ByVal c1&, ByVal c2&, v As Boolean)
Dim r1&, G1&, B1&, r2&, G2&, B2&, B() As Byte
Dim i&, lR!, lG!, lB!, dR!, dG!, dB!, BI&(9), xx&, yy&, dd&, hRPen&
    If dx = 0 Or dy = 0 Then Exit Sub
    If v Then xx = 1: yy = dy: dd = dy Else xx = dx: yy = 1: dd = dx
    SplitRGB c1, r1, G1, B1: SplitRGB c2, r2, G2, B2: ReDim B(dd * 4 - 1)
    dR = (r2 - r1) / (dd - 1): lR = r1: dG = (G2 - G1) / (dd - 1): lG = G1: dB = (B2 - B1) / (dd - 1): lB = B1
    For i = 0 To (dd - 1) * 4 Step 4: B(i + 2) = lR: lR = lR + dR: B(i + 1) = lG: lG = lG + dG: B(i) = lB: lB = lB + dB: Next
    BI(0) = 40: BI(1) = xx: BI(2) = -yy: BI(3) = 2097153: StretchDIBits DC, X, Y, dx, dy, 0, 0, xx, yy, B(0), BI(0), 0, vbSrcCopy
End Sub
Private Sub Form_Load()
'RGB(240, 240, 255)
'RGB(210, 210, 210)
Dim iMth&, iYear&, n&, h&, k&, j&, l&
Dim bSegoe As Boolean, bTahoma As Boolean, bArial As Boolean

    'check fonts
    For n = 0 To Screen.FontCount - 1
        If LCase(Screen.Fonts(n)) = "segoe ui" Then
            bSegoe = True
        ElseIf LCase(Screen.Fonts(n)) = "arial" Then
            bArial = True
        ElseIf LCase(Screen.Fonts(n)) = "tahoma" Then
            bTahoma = True
        End If
    Next

    If bSegoe Then
        Me.FontName = "Segoe UI"
    ElseIf bArial Then
        Me.FontName = "Arial"
    Else
        Me.FontName = "Tahoma"
    End If

    CurrMonth = Month(DateAdd("m", -3, Now)): CurrYear = Year(Now)

    DrawScreen

    h = CreateRoundRectRgn(0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 11, 11)
    SetWindowRgn Me.hwnd, h, False
    
    CurrAlpha = 255
    'SetWindowTopmost Me.hwnd
    
    Me.Move CLng(GetSetting("Widgets", "Calendar", "Left", Screen.Width - (Me.Width * 1.2))), _
                    CLng(GetSetting("Widgets", "Calendar", "Top", Screen.Height - (Me.Height * 1.2)))
    
    If Me.Left > Screen.Width Or (Me.Left + Me.Width) < 0 Or _
    Me.Top > Screen.Height Or (Me.Top + Me.Height) < 0 Then
        Me.Move Screen.Width - (Me.Width * 1.2), Screen.Height - (Me.Height * 1.5)
    End If
    
End Sub

Sub DrawScreen()
Dim k&, j&, l&

    
    Me.Cls

    Gradient Me.hDC, 0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, RGB(210, 210, 210), RGB(240, 240, 255), True
    Gradient Me.hDC, 0, 0, Me.Width \ Screen.TwipsPerPixelX, 25, RGB(240, 240, 255), RGB(190, 190, 190), True

    k = 0: j = CurrMonth: l = CurrYear
    Do
        DrawCal j, l, k
        j = j + 1
        If j = 13 Then
            j = 1
            l = l + 1
        End If
        k = k + 1
    Loop Until k = 7
    
End Sub
Sub DrawCal(iMth&, iYear&, iPos&)
Dim s$, p&, X&, Y&, xOff&, yOff&, n&, h&
    
    xOff = 20 + (iPos * 180): yOff = 20
    
    WriteText Me.hDC, xOff - 6, 3, Format(DateSerial(2007, iMth, 1), "Mmmm") & " " & iYear, 0

    For p = 1 To 7
        WriteText Me.hDC, xOff + 5 + ((p - 1) * 20), 25, Mid$("SMTWTFS", p, 1), 0
    Next
    X = xOff + (Offset(iMth, iYear) * 25)
    Y = 2.25 * yOff
    n = 1 + Offset(iMth, iYear)
    For p = 1 To DaysInMonth(iMth, iYear)
        If iMth = Month(Now) And iYear = Year(Now) And p = Day(Now) Then
            WriteText Me.hDC, X, Y, CStr(p), RGB(200, 0, 0)
        Else
            WriteText Me.hDC, X, Y, CStr(p), 0
        End If
        X = X + 25
        n = n + 1
        If n > 7 Then
            n = 1
            X = xOff
            Y = Y + yOff
        End If
    Next
    
    

End Sub
Private Sub WriteText(hDC&, X&, Y&, s$, fc&)
    SetTextColor Me.hDC, vbWhite
    TextOut Me.hDC, X + 1, Y + 1, s, Len(s)
    SetTextColor Me.hDC, fc
    TextOut hDC, X, Y, s, Len(s)
End Sub
Private Function DaysInMonth&(iM&, iY&)
Dim dteStart As Date
Dim dteEnd As Date
    dteStart = DateSerial(iY, iM, 1)
    dteEnd = DateAdd("m", 1, dteStart)
    DaysInMonth = DateDiff("d", dteStart, dteEnd)
End Function
Private Function Offset&(iM&, iY&)
Dim sDte$
    Offset = 0
    sDte = DateSerial(iY, iM, 1)
    Select Case Format(sDte, "Ddd")
        Case "Sun"
        Offset = 0
        Case "Mon"
        Offset = 1
        Case "Tue"
        Offset = 2
        Case "Wed"
        Offset = 3
        Case "Thu"
        Offset = 4
        Case "Fri"
        Offset = 5
        Case "Sat"
        Offset = 6
    End Select
End Function





Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Widgets", "Calendar", "Left", Me.Left
    SaveSetting "Widgets", "Calendar", "Top", Me.Top
End Sub
