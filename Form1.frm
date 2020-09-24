VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   420
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5460
      Top             =   60
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   60
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4620
      Top             =   60
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4200
      Top             =   60
   End
   Begin VB.PictureBox picScene 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4200
      Left            =   60
      Picture         =   "Form1.frx":0E42
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image imgEgg 
      Height          =   360
      Left            =   120
      Picture         =   "Form1.frx":3E6E4
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMushroom 
      Height          =   360
      Left            =   600
      Picture         =   "Form1.frx":3EDCE
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSpiderLeft 
      Height          =   480
      Index           =   1
      Left            =   3600
      Picture         =   "Form1.frx":3F4B8
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSpiderLeft 
      Height          =   480
      Index           =   0
      Left            =   3120
      Picture         =   "Form1.frx":3FD82
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBG 
      Height          =   480
      Left            =   540
      Picture         =   "Form1.frx":4064C
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFrog 
      Height          =   480
      Index           =   1
      Left            =   1020
      Picture         =   "Form1.frx":40E8E
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFrog 
      Height          =   480
      Index           =   0
      Left            =   1500
      Picture         =   "Form1.frx":41758
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSpiderUp 
      Height          =   480
      Index           =   1
      Left            =   2040
      Picture         =   "Form1.frx":42022
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSpiderUp 
      Height          =   480
      Index           =   0
      Left            =   2580
      Picture         =   "Form1.frx":428EC
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'   writen by yidie
'   http://hi.baidu.com/yi_die
'   xiaocaiyd@sohu.com
'   2007-8-30
'***************************************************

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function PtInRect Lib "user32" _
    (lpRect As RECT, _
    ByVal x As Long, _
    ByVal y As Long) As Long

Private Declare Function timeGetTime Lib "winmm.dll" _
    () As Long

Private Declare Function sndPlaySound Lib "winmm.dll" _
    Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Private Const A_DEGREE As Double = 3.14159265 / 180

Dim x(220) As Integer, y(220) As Integer                            'data for snake
Dim rectFrog As RECT, rectSpider(1 To 100) As RECT                  '
Dim intSnakeLength As Integer, s As Integer, r As Double            'length of snake,direction,move step
Dim intSpider As Integer, intTime As Integer                        'num of spider,delay time
Dim score As Long, hightScore As Long                               'score,hight score
Dim blnIsPause As Boolean, blnIsLast As Boolean
Dim intDelay As Integer, intDelayF As Integer, intIndexF As Integer '
Dim StrSoundsPath As String                                         '

Dim rectMushroom As RECT, intTimeMushroom As Integer                '
Dim blnIsEatMushroom As Boolean                                     '

Dim lngDelaySpider As Long                                          'Ê¹Ö©ÖëÑÓ³Ù³öÏÖ
Dim lngDelayMushroom As Long                                        '³Ôµ½Ä¢¹½ÑÓ³Ù

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(App.Title, Me.Name, "hight", hightScore)       '±£´æ×î¸ß·Ö
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    End
End Sub

Private Sub Form_Load()
    
    Dim i As Integer, j As Integer
    
    If Right(App.Path, 1) <> "\" Then
        StrSoundsPath = App.Path & "\Sounds\"
    Else
        StrSoundsPath = App.Path & "Sounds\"
    End If

    hightScore = GetSetting(App.Title, Form1.Name, "hight", 0)         'get hight score
    
    Form1.PaintPicture picScene.Picture, Form1.ScaleWidth / 2 - 150, Form1.ScaleHeight / 2 - 140
    
    picBG.Move 0, 0, Form1.ScaleWidth, Form1.ScaleHeight
    
    'draw background
    For i = 0 To Form1.ScaleWidth Step imgBG.Width
        For j = 0 To Form1.ScaleHeight Step imgBG.Height
            picBG.PaintPicture imgBG.Picture, i, j
        Next
    Next
    
    picScene.FontBold = True
    PlaySound "music.wav"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then s = s - 5             'move counterclockwise
    If KeyCode = vbKeyRight Then s = s + 5            'move clockwise
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Timer2.Enabled Then Exit Sub
    Select Case KeyCode
        Case vbKeyPause, vbKeyDown          'pause
            If Timer1.Enabled = True Then
                blnIsPause = True
                Timer1.Enabled = False
                Timer3.Enabled = False
                Timer4.Enabled = True
            End If
        Case vbKeyEscape                    'exit
            Unload Me
        Case vbKeyReturn, vbKeyUp           'start\continue
            If blnIsPause Then
                Timer4.Enabled = False
                Timer1.Enabled = True
                Timer3.Enabled = True
                blnIsPause = False
            ElseIf Timer1.Enabled = False And Timer2.Enabled = False Then
                Timer1.Enabled = True
                Timer3.Enabled = True
                InitGame
            End If
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Timer2.Enabled Then Exit Sub
    If Button = vbRightButton Then                      'pause\continue
        If Timer1.Enabled = True Then
            blnIsPause = True
            Timer1.Enabled = False
            Timer3.Enabled = False
            Timer4.Enabled = True
        ElseIf Timer4.Enabled = False Then
            Timer1.Enabled = True
            Timer3.Enabled = True
            InitGame
        Else
            blnIsPause = False
            Timer1.Enabled = True
            Timer3.Enabled = True
            Timer4.Enabled = False
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, iX As Single, iY As Single)
    
    If Button <> vbLeftButton Then Exit Sub
    
    Dim detaX As Double, detaY As Double
    
    detaX = iX - x(0): detaY = iY - y(0)
    
    If detaX = 0 Then
        If detaY < 0 Then s = 0 Else s = 180
    ElseIf detaY = 0 Then
        If detaX < 0 Then s = 270 Else s = 90
    Else
        s = IIf(detaX < 0, 270, 90) + Atn(detaY / detaX) / A_DEGREE
    End If
End Sub

'draw a single frame
Private Sub Timer1_Timer()
    Dim i As Integer, j As Integer, t As Double
    Dim newX As Single, newY As Single
    Dim detaX As Double, detaY As Double
    
    picScene.Cls
    
    If x(0) < 0 Then x(0) = Me.ScaleWidth
    If y(0) < 0 Then y(0) = Me.ScaleHeight
    If x(0) > Me.ScaleWidth Then x(0) = 0
    If y(0) > Me.ScaleHeight Then y(0) = 0
    
    'draw spider
    For i = 1 To intSpider
        If i = intSpider Then
            If blnIsLast Then
                'delay 1500 ms
                If timeGetTime - lngDelaySpider >= 1500 Then
                    blnIsLast = False
                    DrawSpider i, 0
                Else
                    picScene.PaintPicture imgEgg.Picture, rectSpider(i).Left + 4, rectSpider(i).Top + 4
                End If
            Else
                DrawSpider i, 4
            End If
        Else
            DrawSpider i
        End If
    Next
    
    
    
    DrawSnake

    t = s * A_DEGREE
    
    detaX = Sin(t)
    detaY = Cos(t)
            
    'eat spider
    For i = 1 To intSpider - 1
        If i = intSpider - 1 And blnIsLast Then Exit For
        newX = x(0): newY = y(0)
        For j = 1 To 15
            newX = newX + detaX: newY = newY - detaY
            If PtInRect(rectSpider(i), newX, newY) Then
                PlaySound "die.wav"
                Timer2.Enabled = True
                Timer1.Enabled = False
                Timer3.Enabled = False
                Exit Sub
            End If
        Next
    Next
        
    'eat frog
    If PtInRect(rectFrog, x(0) + detaX * 15, y(0) - detaY * 15) Then
        PlaySound
        intDelayF = 0
        intSnakeLength = intSnakeLength + 2
        intSpider = intSpider + 1
        x(intSnakeLength) = x(intSnakeLength - 1)
        y(intSnakeLength) = y(intSnakeLength - 1)
        score = score + intTime * 10& + intSpider * 100&
        intTime = 110 - (intSpider \ 5) * 5
        If hightScore < score Then hightScore = score
        Me.Caption = "Ì°Ê³Éß  ¼ÇÂ¼£º" & hightScore & "  ÄúµÄµÃ·Ö£º" & score
        blnIsLast = True
        
        'finish
        If intSnakeLength >= 220 Then
            Timer1.Enabled = False
            Congration
            Exit Sub
        End If
        
        lngDelaySpider = timeGetTime
        
        If intSpider Mod 10 = 0 Then
        'draw mushroom
            With rectMushroom
                .Left = Int(Rnd * (Me.ScaleWidth - 40)) + 20
                .Top = Int(Rnd * (Me.ScaleHeight - 40)) + 20
                .Right = .Left + 20
                .Bottom = .Top + 20
            End With
            intTimeMushroom = 5
        End If
               
        With rectFrog
            'set spider's location
            rectSpider(intSpider).Left = .Left
            rectSpider(intSpider).Top = .Top
            rectSpider(intSpider).Right = .Right
            rectSpider(intSpider).Bottom = .Bottom
            
            'next frog's location
            .Left = Fix(Rnd * (Me.ScaleWidth - 40)) + 20
            .Top = Fix(Rnd * (Me.ScaleHeight - 40)) + 20
            .Right = .Left + 32
            .Bottom = .Top + 32
        End With
    End If
    
    'eat mushroom
    If PtInRect(rectMushroom, x(0) + detaX * 15, y(0) - detaY * 15) Then
        PlaySound "fun.wav"
        blnIsEatMushroom = True
        intTimeMushroom = 0
        lngDelayMushroom = timeGetTime
        'move mushroom out
        rectMushroom.Left = -100
        rectMushroom.Top = -100
        rectMushroom.Right = -100
        rectMushroom.Bottom = -100
        
        score = score + 10000&
        If hightScore < score Then hightScore = score
        Me.Caption = "Snake  hight score£º" & hightScore & "  your score£º" & score
        
    End If
    
    'draw mushroom
    If intTimeMushroom > 0 Then
        With rectMushroom
            picScene.PaintPicture imgMushroom.Picture, .Left, .Top
            picScene.FontSize = 10
            picScene.ForeColor = vbYellow
            picScene.CurrentX = .Left + 4
            picScene.CurrentY = .Top - picScene.TextHeight("A")
            picScene.Print intTimeMushroom
        End With
    End If
        
    DrawFrog
        
    'trun a frame
    Me.Picture = picScene.Image
End Sub

'over
Private Sub Timer2_Timer()
    Delay 500
    PlaySound "bom.wav"
    Dim i As Integer, j As Integer, t As Double
    For i = 25 To 1 Step -2
        picScene.Cls
        For j = 0 To 360 Step 30
            t = j * A_DEGREE
            picScene.Line (x(0) + Sin(t) * i, y(0) - Cos(t) * i)-(x(0) + Sin(t) * (i + 5), y(0) - Cos(t) * (i + 5)), vbRed
        Next
        Me.Picture = picScene.Image
        Delay 100
    Next
    picScene.Cls
    picScene.FontSize = 120
    picScene.ForeColor = vbRed
    picScene.CurrentX = (Me.ScaleWidth - picScene.TextWidth("you lose")) / 2
    picScene.CurrentY = (Me.ScaleHeight - picScene.TextHeight("you lose")) / 2
    picScene.Print "you lose"
    
    Me.Picture = picScene.Image
        
    Delay 1000
    
    If MsgBox("score£º" & score & vbCrLf & "play again£¿", vbYesNo + vbInformation, "snake") = vbYes Then
        Timer1.Enabled = True
        Timer3.Enabled = True
        InitGame
    Else
        Unload Me
    End If
    
    Timer2.Enabled = False

End Sub

'begin to count down
Private Sub Timer3_Timer()
    intTime = intTime - 1
    If intTime = 0 Then
        Timer2.Enabled = True
        Timer1.Enabled = False
        Timer3.Enabled = False
    End If
    If intTimeMushroom <> 0 Then
        intTimeMushroom = intTimeMushroom - 1
        If intTimeMushroom = 0 Then
            blnIsEatMushroom = False
        End If
    End If
End Sub

'if pause
Private Sub Timer4_Timer()
    score = score - 10
    Me.Caption = "Snake  hight score£º" & hightScore & "  your score£º" & score
End Sub

Private Sub InitGame()
    
    Dim i As Integer
    
    picScene.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    picScene.Picture = picBG.Image
    picScene.Cls
    picScene.FontName = "MS Sans Serif"
    intSnakeLength = 20
    'data for snake
    For i = 0 To intSnakeLength
        x(i) = Me.ScaleWidth \ 2
        y(i) = Me.ScaleHeight \ 2
    Next
    
    r = 7
    s = 0
    
    Randomize Timer
    'frog
    With rectFrog
        .Left = Int(Rnd * (Me.ScaleWidth - 40)) + 20
        .Top = Int(Rnd * (Me.ScaleHeight - 40)) + 20
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
    
    intSpider = 0
    intTime = 110 - (intSpider \ 5) * 5
    
    intTimeMushroom = 0
    blnIsEatMushroom = False
    
    score = 0
    Me.Caption = "Snake  hight score£º" & hightScore & "  your score£º" & score
End Sub

'success
Private Sub Congration()
    PlaySound "win.wav"

    picScene.FontSize = 120
    picScene.ForeColor = vbRed
    picScene.CurrentX = (Me.ScaleWidth - picScene.TextWidth("Congration")) / 2
    picScene.CurrentY = (Me.ScaleHeight - picScene.TextHeight("Congration")) / 2
    picScene.Print "Congration"
    
    Me.Picture = picScene.Image
    
    Delay 10000
    
    If MsgBox("score£º" & score & vbCrLf & "play again£¿", vbYesNo + vbInformation, "snake") = vbYes Then
        Timer1.Enabled = True
        Timer3.Enabled = True
        InitGame
    Else
        Unload Me
    End If
End Sub


Private Sub PlaySound(Optional ByVal fileName As String = "ching.wav")
    sndPlaySound StrSoundsPath & fileName, 3&
End Sub


Private Sub DrawSpider(ByVal i As Integer, Optional ByVal step As Integer = 6)
    
    If blnIsEatMushroom Then
        'spider stop move ,delay 4 seconds
        If timeGetTime - lngDelayMushroom >= 4000 Then blnIsEatMushroom = False
        step = 0
    End If
    
    With rectSpider(i)
        Select Case (i And 3)
            Case 0  'move up
                .Top = .Top - Fix(Rnd * step)
                If .Top < -32 Then .Top = Form1.ScaleHeight
                .Bottom = .Top + 32
                picScene.PaintPicture imgSpiderUp(Fix(2 * Rnd)).Picture, .Left, .Top
            Case 1  'move down
                .Top = .Top + Fix(Rnd * step)
                If .Top > Form1.ScaleHeight Then .Top = -32
                .Bottom = .Top + 32
                picScene.PaintPicture imgSpiderUp(Fix(2 * Rnd)).Picture, .Left, .Top, 32, 32, 0, 31, 32, -32
            Case 2  'move left
                .Left = .Left - Fix(Rnd * step)
                If .Left < -32 Then .Left = Form1.ScaleWidth
                .Right = .Left + 32
                picScene.PaintPicture imgSpiderLeft(Fix(2 * Rnd)).Picture, .Left, .Top
            Case 3  'move right
                .Left = .Left + Fix(Rnd * step)
                If .Left > Form1.ScaleWidth Then .Left = -32
                .Right = .Left + 32
                picScene.PaintPicture imgSpiderLeft(Fix(2 * Rnd)).Picture, .Left, .Top, 32, 32, 31, 0, -32, 32
        End Select
    End With
End Sub

Private Sub DrawFrog(Optional ByVal step As Integer = 100)
    
    If blnIsEatMushroom Then step = 0 'stop move
    
    With rectFrog
        picScene.FillColor = RGB(0, 255, 0)
        intDelayF = (intDelayF + 1) And 31
        If intDelayF = 31 Then  'move to new place
            intIndexF = 1 - intIndexF
            .Left = .Left + Fix(step * (0.5 - Rnd))
            .Top = .Top + Fix(step * (0.5 - Rnd))
            If .Left < -16 Then .Left = Form1.ScaleWidth - 16
            If .Left > Form1.ScaleWidth Then .Left = -16
            If .Top < -16 Then .Top = Form1.ScaleHeight - 16
            If .Top > Form1.ScaleHeight Then .Top = -16
            .Bottom = .Top + 32
            .Right = .Left + 32
            
            If step <> 0 Then PlaySound "pop.wav"
        End If
        'draw frog
        picScene.PaintPicture imgFrog(intIndexF).Picture, .Left, .Top
        picScene.FontSize = 10
        picScene.ForeColor = vbWhite
        picScene.CurrentX = .Left
        picScene.CurrentY = .Top - picScene.TextHeight("A")
        picScene.Print intTime
    End With

End Sub

'»­Éß
Private Sub DrawSnake()
    Dim i As Integer
    Dim t As Double, t2 As Double, detaX As Double, detaY As Double
    
    'snake
    For i = intSnakeLength To 1 Step -1
        x(i) = x(i - 1)
        y(i) = y(i - 1)
        
        picScene.DrawWidth = Int((intSnakeLength - i) / intSnakeLength * 11 + 2)
        
        If i <> intSnakeLength Then
           If Abs(x(i) - x(i + 1)) > 10 Or Abs(y(i) - y(i + 1)) > 10 Then
           Else
               picScene.Line (x(i), y(i))-(x(i + 1), y(i + 1)), IIf((i And 1), RGB(&H0, &H55, &H0), RGB(&HFF, &HCC, 0))
           End If
        End If
    Next

    t = s * A_DEGREE
    t2 = 45 * A_DEGREE
    
    detaX = Sin(t)
    detaY = Cos(t)

    x(0) = x(0) + detaX * r
    y(0) = y(0) - detaY * r
    
    If Abs(x(0) - x(1)) > 10 Or Abs(y(0) - y(1)) > 10 Then
    Else
        picScene.Line (x(0), y(0))-(x(1), y(1)), RGB(&HFF, &HCC, 0)
    End If
    
    'head
    picScene.DrawWidth = 1
    picScene.FillColor = RGB(&H0, &H55, &H0)
    picScene.Circle (x(0), y(0)), 8, RGB(&H0, &H55, &H0)
    picScene.Circle (x(0) + detaX * 7, y(0) - detaY * 7), 4, RGB(&H0, &H55, &H0)
    'eyes
    picScene.FillColor = vbRed
    picScene.Circle (x(0) + Sin(t + t2) * 8, y(0) - Cos(t + t2) * 8), 1, vbRed
    picScene.Circle (x(0) + Sin(t - t2) * 8, y(0) - Cos(t - t2) * 8), 1, vbRed
    
    If Int(Rnd * 3) = 0 Then
      'tongue
      t2 = 5 * A_DEGREE
      
      picScene.Line (x(0) + detaX * 11, y(0) - detaY * 11)-(x(0) + detaX * 15, y(0) - detaY * 15), vbRed
      picScene.Line (x(0) + Sin(t + t2) * 22, y(0) - Cos(t + t2) * 22)-(x(0) + detaX * 15, y(0) - detaY * 15), vbRed
      picScene.Line (x(0) + Sin(t - t2) * 22, y(0) - Cos(t - t2) * 22)-(x(0) + detaX * 15, y(0) - detaY * 15), vbRed
    
    End If

End Sub

'delay,ms
Private Sub Delay(ByVal n As Long)
    Dim lngT As Long
    lngT = timeGetTime
    While (timeGetTime - lngT < n)
        DoEvents
    Wend
End Sub

