VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H0099AA99&
   BorderStyle     =   0  'None
   Caption         =   "Animate Desktopo Icons"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":1042
   ScaleHeight     =   3210
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox rotText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   25
      Top             =   2085
      Width           =   495
   End
   Begin VB.HScrollBar rotScroll 
      Height          =   235
      LargeChange     =   2
      Left            =   1920
      Max             =   5
      Min             =   1
      TabIndex        =   24
      Top             =   2085
      Value           =   1
      Width           =   2300
   End
   Begin VB.Timer stat_timer 
      Interval        =   1000
      Left            =   2760
      Top             =   3360
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H003399FF&
      Caption         =   "Run when windows starts"
      Height          =   255
      Left            =   1830
      MaskColor       =   &H008080FF&
      TabIndex        =   21
      ToolTipText     =   "Check or uncheck to start this program at windows startup"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox speedText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.HScrollBar speedScroll 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   15
      Min             =   1
      TabIndex        =   17
      Top             =   1800
      Value           =   1
      Width           =   2300
   End
   Begin animate_icons.Transparent Transparent1 
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      MaskColor       =   16711935
   End
   Begin VB.HScrollBar yRadSpin 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   300
      TabIndex        =   14
      Top             =   1515
      Width           =   2300
   End
   Begin VB.HScrollBar xradSpin 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   400
      TabIndex        =   13
      Top             =   1230
      Width           =   2300
   End
   Begin VB.HScrollBar angleSpin 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   720
      TabIndex        =   12
      Top             =   945
      Width           =   2300
   End
   Begin VB.HScrollBar topSpin 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   600
      TabIndex        =   11
      Top             =   660
      Width           =   2300
   End
   Begin VB.HScrollBar leftSpin 
      Height          =   235
      LargeChange     =   5
      Left            =   1920
      Max             =   800
      TabIndex        =   10
      Top             =   375
      Width           =   2300
   End
   Begin VB.TextBox angleText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   8
      Top             =   945
      Width           =   495
   End
   Begin VB.TextBox topText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   6
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox leftText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   375
      Width           =   495
   End
   Begin VB.TextBox yradText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   1515
      Width           =   495
   End
   Begin VB.TextBox xradText 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   0
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Times Rotation"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   26
      Top             =   2085
      Width           =   1155
   End
   Begin VB.Label inf 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   4125
      TabIndex        =   23
      ToolTipText     =   "Close this program"
      Top             =   2490
      Width           =   255
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1605
      TabIndex        =   22
      Top             =   2850
      Width           =   2085
   End
   Begin VB.Image scdown 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1410
      Picture         =   "Form1.frx":21E4
      Top             =   3645
      Width           =   1245
   End
   Begin VB.Image aniup 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":256D
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Image save 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   285
      Picture         =   "Form1.frx":2958
      ToolTipText     =   "Save Settings"
      Top             =   2805
      Width           =   1245
   End
   Begin VB.Image scup 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1410
      Picture         =   "Form1.frx":2D3F
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Image anidown 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":3126
      Top             =   3645
      Width           =   1245
   End
   Begin VB.Image animate 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   285
      Picture         =   "Form1.frx":34AD
      ToolTipText     =   "Animate Icons"
      Top             =   2535
      Width           =   1245
   End
   Begin VB.Label about 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   3795
      TabIndex        =   20
      ToolTipText     =   "About the Author"
      Top             =   2835
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   675
      TabIndex        =   19
      Top             =   1815
      Width           =   555
   End
   Begin VB.Label extBut 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   4140
      TabIndex        =   16
      ToolTipText     =   "Close this program"
      Top             =   2820
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Angle"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   375
      TabIndex        =   9
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   855
      TabIndex        =   7
      Top             =   660
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   855
      TabIndex        =   5
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Y-Radius"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   525
      TabIndex        =   3
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X-Radius"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   510
      TabIndex        =   1
      Top             =   1230
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageByLong& Lib "user32" Alias _
"SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" _
(ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName _
As String, ByVal lpWindowName As String)
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Const LVM_GETTITEMCOUNT& = (&H1000 + 4)
Private Const LVM_SETITEMPOSITION& = (&H1000 + 15)

Private Const PI = 3.1415926535898
Dim DegreeToInc As Double
Dim hdesk&, i&, icount&, X&, Y&
Dim xrad As String, yrad, fromLeft, fromTop, pSpeed, PiniAngle, PiniAngleTemp, numOfRotation
Dim moving As Boolean, makeChangesToRegistry, settingsAltered, settings_saved
Dim lastX As Single, lastY As Single
Dim appName As String
Dim entryFound As Boolean ' If the program has registry entry
Private Type a4
    a As String * 4
End Type
Private Type l4
    l As Long
End Type

Public Sub MoveIcons()
hdesk = FindWindow("progman", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "shelldll_defview", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "syslistview32", vbNullString)
'hdesk is the handle of the Desktop's syslistview32

icount = SendMessageByLong(hdesk, LVM_GETTITEMCOUNT, 0, 0)
DegreeToInc = CInt(360 / icount)
'0 is "My Computer"
For i = 0 To icount - 1
X = CInt(fromLeft) - CInt(xrad) * Cos((i * DegreeToInc + CInt(PiniAngle)) * PI / 180)
Y = CInt(fromTop) + CInt(yrad) * Sin((i * DegreeToInc + CInt(PiniAngle)) * PI / 180) 'set the position parameters in pixel
'The wParam must be i
Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, CLng(X + Y * &H10000))
Next
End Sub

Private Sub about_Click()
    MsgBox vbCrLf & _
           "© " & vbCrLf & _
           "  © " & vbCrLf & _
           "©   Program created by Mahbubur Rahman " & vbCrLf & _
           "©   pappu@inbox.net >> www.pappu.tk " & vbCrLf & _
           "  ©" & vbCrLf & _
           "©"
End Sub

Private Sub animate_Click()
    animateIcons
End Sub


Private Sub animateIcons()
    Dim i As Integer
    PiniAngleTemp = PiniAngle
    For i = CInt(PiniAngle) To CInt(PiniAngle) + 360 * CInt(numOfRotation) Step CInt(pSpeed)
        PiniAngle = "" & i
        Call MoveIcons
        DoEvents
    Next
    PiniAngle = PiniAngleTemp
End Sub


Private Sub animate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    animate.Picture = anidown.Picture
End Sub
Private Sub animate_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    animate.Picture = aniup.Picture
End Sub

Private Sub Check1_Click()
    makeChangesToRegistry = True: settingsAltered = True
End Sub

Private Sub extBut_Click()
    If settingsAltered = True And settings_saved = False Then
        If MsgBox("[ Settings have been Altered ]" & vbCrLf & vbCrLf & "You wanna save changes ?", vbYesNo, "Save Changes") = vbYes Then
            Call save_settings(True)
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Function check_registry(myKey As String)
    Dim lRet As Long, hkey As Long
    lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, hkey)
    If lRet Then MsgBox "Error accessing HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run": Exit Function
    
    Dim lIndex As Long, aVName$, lVName As Long, lType As Long, aData$, lData As Long
    Dim aAdd$, l As Long, a4 As a4, l4 As l4, a$
    lVName = 100                ' Name buffer length
    aVName$ = Space$(lVName)    ' Name buffer
    lData = 100                 ' Data buffer length
    aData$ = Space$(lData)      ' Data buffer
    lRet = RegEnumValue(hkey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Do Until lRet = ERROR_NO_MORE_ITEMS
        aAdd$ = ""
        If Left$(aVName$, lVName) = myKey Then: entryFound = True

        'List2.AddItem aAdd$
        lVName = 100        ' You MUST reset these buffer lengths, because the RegEnumValue call
        lData = 100         ' changed them to = # of bytes copied
        lIndex = lIndex + 1
        lRet = RegEnumValue(hkey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Loop
    lRet = RegCloseKey(hkey)
    check_registry = entryFound
End Function

Private Sub Form_Initialize()
    Dim screenH As Long, screenW
    appName = App.EXEName & ".exe"
    screenH = Screen.Height / Screen.TwipsPerPixelY
    screenW = Screen.Width / Screen.TwipsPerPixelX
    leftSpin.Max = screenW: topSpin.Max = screenH
    Dim p As New Cini
    p.FileName = App.Path & "\icon_settings.ini"
    p.ApplicationKey = "icons_geometry"
    xrad = p.GetValue("xradius", "250")
    xradText.Text = xrad: xradSpin.Value = xrad
    yrad = p.GetValue("yradius", "90")
    yradText.Text = yrad: yRadSpin.Value = yrad
    fromLeft = p.GetValue("left", "250")
    leftText.Text = fromLeft: leftSpin.Value = fromLeft
    fromTop = p.GetValue("top", "100")
    topText.Text = fromTop: topSpin.Value = fromTop
    PiniAngle = p.GetValue("iniAngle", "0")
    angleText.Text = PiniAngle: angleSpin.Value = PiniAngle
    pSpeed = p.GetValue("speed", "1")
    speedText.Text = pSpeed: speedScroll.Value = pSpeed
    numOfRotation = p.GetValue("numRotation", "1")
    rotText.Text = numOfRotation: rotScroll.Value = numOfRotation
    If check_registry("aniIcons") = True Then
        Check1.Value = 1
        entryFound = True
    Else
        Check1.Value = 0
        entryFound = False
    End If
    Set p = Nothing
    settingsAltered = False: settings_saved = False
    If LCase(Command) = "aniandexit" Then
    Call animateIcons
    End
    End If
End Sub

Private Sub Form_Load()
If LCase(Command) = "aniandexit" Then
    Form1.Visible = False
Else
    Show
    Set Transparent1.MaskPicture = Form1.Picture
    speedScroll.SetFocus
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = True
    lastX = X
    lastY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If moving Then
        Me.Move (Me.Left + X - lastX), (Me.Top + Y - lastY)
        DoEvents
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = False
End Sub

Private Sub save_settings(unloadAfterSave As Boolean)
    Dim p As New Cini
    Dim pRet As Long
    p.FileName = App.Path & "\icon_settings.ini"
    If GetFileAttributes(p.FileName) <> 32 Then
        pRet = SetFileAttributes(p.FileName, 32)    ' Removing Read only attributes of the ini file if some one has
    End If                                          ' mistakenly made that read-only and to allow save settings
    p.ApplicationKey = "icons_geometry"
    p.SetValue "xradius", xrad
    p.SetValue "yradius", yrad
    p.SetValue "left", fromLeft
    p.SetValue "top", fromTop
    p.SetValue "iniAngle", PiniAngle
    p.SetValue "speed", pSpeed
    p.SetValue "numRotation", numOfRotation
    Set p = Nothing
    If makeChangesToRegistry = True Then
        If Check1.Value = 1 And entryFound = False Then
            Dim Zero, hkey As Long
            RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", Zero, KEY_ALL_ACCESS, hkey
            SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "aniIcons", appName & " aniandexit", REG_SZ
        End If
        If Check1.Value = 0 And entryFound = True Then
            DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "aniIcons"
        End If
    End If
    showStatus "Settings have been saved", 1000
    settingsAltered = False
    If unloadAfterSave = True Then: Unload Me
End Sub

Private Sub showStatus(pText As String, pInterval As Long)
    status.Caption = pText
    stat_timer.Interval = pInterval
    stat_timer.Enabled = True
End Sub

Private Sub inf_Click()
    MsgBox "To run the program at system startup --" & vbCrLf & vbCrLf & _
    "-> Copy the ""animate_icons.exe"" and ""icon_settings.inf"" into your windows directory" & vbCrLf & _
    "-> Execute the ""animate_icons.exe"" file from within the windows directory" & vbCrLf & vbCrLf & _
    "To Restore your desktop icons --" & vbCrLf & vbCrLf & _
    "->I guess u no. If u don't, right click on desktop >>Arrange Icons by >> Name or Type or Modified."
End Sub

Private Sub leftSpin_scroll()
    leftText.Text = leftSpin.Value: fromLeft = leftSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub leftSpin_change()
    leftText.Text = leftSpin.Value: fromLeft = leftSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub save_Click()
    save_settings False
End Sub

Private Sub save_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    save.Picture = scup.Picture
End Sub
Private Sub save_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    save.Picture = scdown.Picture
End Sub

Private Sub speedScroll_scroll()
    speedText.Text = speedScroll.Value: pSpeed = speedScroll.Value: settingsAltered = True
End Sub
Private Sub speedScroll_change()
    speedText.Text = speedScroll.Value: pSpeed = speedScroll.Value: settingsAltered = True
End Sub
Private Sub stat_timer_Timer()
    status.Caption = ""
    stat_timer.Enabled = False
End Sub
Private Sub topSpin_scroll()
    topText.Text = topSpin.Value: fromTop = topSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub topSpin_change()
    topText.Text = topSpin.Value: fromTop = topSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub xradSpin_Change()
    xradText.Text = xradSpin.Value: xrad = xradSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub xradSpin_scroll()
    xradText.Text = xradSpin.Value: xrad = xradSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub yradSpin_Change()
    yradText.Text = yRadSpin.Value: yrad = yRadSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub
Private Sub yradSpin_scroll()
    yradText.Text = yRadSpin.Value: yrad = yRadSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub angleSpin_Change()
    angleText.Text = angleSpin.Value: PiniAngle = angleSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub

Private Sub angleSpin_scroll()
    angleText.Text = angleSpin.Value: PiniAngle = angleSpin.Value: settingsAltered = True
    Call MoveIcons
End Sub
Private Sub rotScroll_Change()
    rotText.Text = rotScroll.Value: numOfRotation = rotScroll.Value: settingsAltered = True
End Sub

Private Sub rotScroll_scroll()
    rotText.Text = rotScroll.Value: numOfRotation = rotScroll.Value: settingsAltered = True
End Sub
