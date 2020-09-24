VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This code is write by chinese
'Sorry my English is poor
Dim dx As New DirectX7
Dim dd As DirectDraw7
Dim blacksurf As DirectDrawSurface7
Dim spritesurf As DirectDrawSurface7
Dim mapsurf As DirectDrawSurface7
Dim startsurf As DirectDrawSurface7 'logo
Dim primary As DirectDrawSurface7
Dim backbuffer As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2 'tank
Dim ddsd4 As DDSURFACEDESC2
Dim ddsd5 As DDSURFACEDESC2 'map
Dim ddsd6 As DDSURFACEDESC2 'logo
Dim brunning As Boolean
Dim binit As Boolean
Dim CurModeActiveStatus As Boolean
Dim bRestore As Boolean
Dim sMedia As String
'###############################
Dim timerid As Long
'###############################
Sub Init()
    '############
    
    '############
    'On Local Error GoTo errOut
            
    Dim file As String
    
    Set dd = dx.DirectDrawCreate("")
    Me.Show
    
    'indicate that we dont need to change display depth
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    
    dd.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
    

            
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
        
    Set primary = dd.CreateSurface(ddsd1)
    
    
    
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    
    backbuffer.GetSurfaceDesc ddsd4
    
    

    'We create a DrawableSurface class from our backbuffer
    'that makes it easy to draw text
    backbuffer.SetForeColor vbGreen
    backbuffer.SetFontTransparency True
    
    ' init the surfaces
    InitSurfaces
                                                    
    binit = True
    brunning = True
    'timerid = timeSetEvent(10, 0, AddressOf fight, 1, 1)
    Do While brunning
    
        fight
        DoEvents
    Loop
    
    
errOut:
    
   EndIT
    
End Sub

Sub InitSurfaces()

    Set blacksurf = Nothing
    Set spritesurf = Nothing
    Set mapsurf = Nothing
    'sMedia = FindMediaDir("lake.bmp")
    'If sMedia = vbNullString Then sMedia = AddDirSep(CurDir)
    
    'load the bitmap into the second surface same size
    'as our back buffer
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = ddsd4.lWidth
    ddsd2.lHeight = ddsd4.lHeight
    Set blacksurf = dd.CreateSurfaceFromFile("back.bmp", ddsd2)
                                                                        
    'load the bitmap into the second surface
    ddsd3.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd3.lWidth = 160
    ddsd3.lHeight = 234
    Set spritesurf = dd.CreateSurfaceFromFile("tank.bmp", ddsd3)
    ddsd5.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd5.lWidth = 40
    ddsd5.lHeight = 568
    Set mapsurf = dd.CreateSurfaceFromFile("map\map.bmp", ddsd5)
    'use black for transparent color key
    ddsd6.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd6.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd6.lWidth = ddsd4.lWidth - 100
    ddsd6.lHeight = ddsd4.lHeight - 80
    Set startsurf = dd.CreateSurfaceFromFile("map\start.bmp", ddsd6)
    Dim key As DDCOLORKEY
    key.low = 0
    key.high = 0
    spritesurf.SetColorKey DDCKEY_SRCBLT, key
    mapsurf.SetColorKey DDCKEY_SRCBLT, key
    backbuffer.SetColorKey DDCKEY_SRCBLT, key

End Sub


Public Sub blt()
Dim i As Long, j As Long
    'On Local Error GoTo errOut
    If binit = False Then Exit Sub
    
    Dim rSprite As RECT
    Dim rSprite2 As RECT
    Dim rPrim As RECT
    
    ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    
    ' if we lost and got back the surfaces, then restore them
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces ' must init the surfaces again if they we're lost
    End If
    'get the rectangle for our source sprite
    rSprite.Bottom = ddsd3.lHeight / 6
    rSprite.Right = ddsd3.lWidth * (26 / 114)

    'paint the background onto our back buffer
    Dim rback As RECT
    rback.Bottom = ddsd2.lHeight
    rback.Right = ddsd2.lWidth
    'rback.Bottom = ddsd4.lHeight
    'rback.Right = ddsd4.lWidth
    Call backbuffer.BltFast(0, 0, blacksurf, rback, DDBLTFAST_WAIT)
    Call backbuffer.drawbox(0, 0, 520, 480)
    Call backbuffer.drawbox(520, 0, 640, 480)
    'blt to the backbuffer from our  surface
    Call drawtank
     For i = 1 To 10
        For j = 1 To 3
            If bomb(i, j).type <> 0 Then Call drawbomb(i, j)
        Next j
    Next i
    If box.shape > 0 Then drawbox
    If Shuaiwudi > 0 Then Drawshuaiwudi
    If start = 0 Then Drawstart
    If passav <= 100 And start = 1 Then Drawpassav 'ÓÎÏ·¿ªÊ¼ºó£¬¹ý¹Øºó×Ô¶¯½øÈëÏÂÒ»¹Ø£¬·ñÔòÊÖ¶¯
    Call drawtext
    'flip the backbuffer to the screen
    primary.Flip Nothing, DDFLIP_WAIT
Exit Sub
errOut:
    EndIT
End Sub

Sub drawtank()
Dim i
For i = 1 To 10
    If tank(i).shape <> 0 Then
        If tank(i).y > 480 - maph Then tank(i).y = 480 - maph
        Call backbuffer.BltFast(tank(i).x - mapw, tank(i).y - maph, spritesurf, tank(i).r, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        If tank(i).wudi > 0 Then drawwudi (i)
    End If
Next i


End Sub
Sub drawbox()
Dim rr As RECT
rr.Left = 0
rr.Top = box.shape * maph * 2 + maph * 9
rr.Right = mapw * 2
rr.Bottom = box.shape * maph * 2 + maph * 9 + maph * 2
Call backbuffer.BltFast(box.x - mapw, box.y - maph, mapsurf, rr, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
box.timer = box.timer - 1
If box.timer = 0 Then box.shape = 0
End Sub
Sub drawtext()
Dim str As String
If passav <= 100 Or start = 0 Then Exit Sub
If thanks(1) = "" Then
    If bug(1) <> "" And bug(2) <> "" Then
        Call backbuffer.SetFillColor(RGB(100, 100, 100))
        Call backbuffer.SetFillStyle(0)
        Call backbuffer.drawbox(100, 200, 420, 280)
        Call backbuffer.drawtext(230, 220, bug(1), False)
        Call backbuffer.drawtext(120, 250, bug(2), False)
        Call backbuffer.SetFillStyle(1)
    End If
Else
    Call backbuffer.SetFillColor(RGB(100, 150, 200))
    Call backbuffer.drawbox(100, 150, 430, 280)
    Call backbuffer.drawtext(110, 160, thanks(1), False)
    Call backbuffer.drawtext(110, 190, thanks(2), False)
    Call backbuffer.drawtext(110, 220, thanks(3), False)
    Call backbuffer.drawtext(110, 250, thanks(4), False)
    Call backbuffer.SetFillStyle(1)
End If
Call backbuffer.drawtext(525, 10, "computer", False)
str = "Tank ¡Á" & enemyleft
Call backbuffer.drawtext(550, 40, str, False)
str = "Round  " & round
Call backbuffer.drawtext(525, 200, str, False)
str = "Player1"
Call backbuffer.drawtext(525, 300, str, False)
str = "Tank ¡Á" & tank(1).life
Call backbuffer.drawtext(550, 330, str, False)
Call backbuffer.drawtext(550, 350, score(1), False)
str = "Player2"
Call backbuffer.drawtext(525, 380, str, False)
str = "Tank ¡Á" & tank(2).life
Call backbuffer.drawtext(550, 410, str, False)
Call backbuffer.drawtext(550, 430, score(2), False)
End Sub
Sub Drawpassav()
Const p = 100
If passav > p Then
    Exit Sub
End If
If start = 0 Then Exit Sub
Dim i As Long
Dim h As Long
Dim ss As String
If round = 0 Then
    
    ss = "Round    1"
Else
    ss = "Round    " & round
End If
Dim rr As RECT
Call backbuffer.SetFillColor(RGB(100, 100, 100))
Call backbuffer.SetFillStyle(0)
Call backbuffer.SetForeColor(RGB(100, 100, 100))
If passav <= p / 4 Then
    h = (passav) * 240 / (p / 4)
        Call backbuffer.drawbox(0, 0, 640, h)
        Call backbuffer.drawbox(0, 480 - h, 640, 480)
ElseIf passav >= p / 4 And passav <= 3 * p / 4 Then
        Call backbuffer.drawbox(0, 0, 640, 480)
        Call backbuffer.SetForeColor(vbGreen)
        Call backbuffer.drawtext(300, 230, ss, False)
Else
    h = (p - passav) * 240 / (p / 4)
        Call backbuffer.drawbox(0, 0, 640, h)
        Call backbuffer.drawbox(0, 480 - h, 640, 480)
End If
passav = passav + 1
Call backbuffer.SetFillStyle(1)
Call backbuffer.SetForeColor(vbGreen)
End Sub
Sub Drawstart()
Dim rr As RECT
rr.Left = 0
rr.Top = 0
rr.Bottom = ddsd6.lHeight
rr.Right = ddsd6.lWidth
Call backbuffer.BltFast(50, 40, startsurf, rr, DDBLTFAST_WAIT)

End Sub
Sub drawshuai(life As Long)
Dim rr As RECT
Dim x As Single
Dim y As Single
x = map(13, 22).x
y = map(13, 22).y
If life = 1 Then
    rr.Left = 0
    rr.Top = 3 * maph
    rr.Right = 2 * mapw
    rr.Bottom = rr.Top + maph * 2
Else
    rr.Left = 0
    rr.Top = 5 * maph
    rr.Right = 2 * mapw
    rr.Bottom = rr.Top + maph * 2
End If
 Call blacksurf.BltFast(x, y, mapsurf, rr, DDBLTFAST_WAIT)

End Sub
Sub drawwudi(index As Long)
If tank(index).shape = 0 Then Exit Sub
Dim oldcol As Long
oldcol = backbuffer.GetForeColor
Call backbuffer.SetForeColor(vbWhite)
If tank(index).wudi Mod 5 < 3 Then
Call backbuffer.DrawRoundedBox(tank(index).x - mapw, tank(index).y - maph - 3, tank(index).x + mapw, tank(index).y + maph, 10, 10)
End If
Call backbuffer.SetForeColor(oldcol)
If tank(index).wudi > 0 Then tank(index).wudi = tank(index).wudi - 1
End Sub
Sub Drawshuaiwudi()
Dim i As Long
Static kx(1 To 8) As Long, ky(1 To 8) As Long

If Shuaiwudi >= 1490 Then
    For i = 1 To 8
        ky(i) = Val(Mid("2322212121212223", i * 2 - 1, 2))
        kx(i) = Val(Mid("1212121314151515", i * 2 - 1, 2))
        map(kx(i), ky(i)).shape = 2
        Call drawmap(map(kx(i), ky(i)).x, map(kx(i), ky(i)).y, map(kx(i), ky(i)).shape)
    Next i
End If
If Shuaiwudi <= 300 And Shuaiwudi >= 2 Then
    If (Shuaiwudi Mod 30) = 1 Then
        For i = 1 To 8
            Call drawmap(map(kx(i), ky(i)).x, map(kx(i), ky(i)).y, 1)
        Next i
    ElseIf (Shuaiwudi Mod 30) = 15 Then
        For i = 1 To 8
            Call drawmap(map(kx(i), ky(i)).x, map(kx(i), ky(i)).y, 2)
        Next i
    End If
End If
If Shuaiwudi = 1 Then
    For i = 1 To 8
        map(kx(i), ky(i)).shape = 1
        Call drawmap(map(kx(i), ky(i)).x, map(kx(i), ky(i)).y, 1)
    Next i
End If
Shuaiwudi = Shuaiwudi - 1
End Sub
Sub drawmap(x As Single, y As Single, shape As Long)
Dim rr As RECT
rr.Left = (shape \ 3) * mapw
rr.Top = (shape Mod 3) * maph
rr.Right = rr.Left + mapw
rr.Bottom = rr.Top + maph
Call blacksurf.BltFast(x, y, mapsurf, rr, DDBLTFAST_WAIT)
        
End Sub

Sub drawbomb(i As Long, j As Long)
Dim oldcol As Long
Dim rr As RECT
If bomb(i, j).type <= 9 Then
    If tank(i).gun = 0 Then
        oldcol = backbuffer.GetForeColor
        Call backbuffer.SetForeColor(vbWhite)
        Call backbuffer.SetFillColor(vbWhite)
        Call backbuffer.DrawCircle(bomb(i, j).x, bomb(i, j).y, 4)
        Call backbuffer.SetForeColor(oldcol)
        Call backbuffer.SetFillStyle(1)
    Else
        oldcol = backbuffer.GetForeColor
        Call backbuffer.SetForeColor(vbBlue)
        Call backbuffer.SetFillColor(vbWhite)
        Call backbuffer.DrawCircle(bomb(i, j).x, bomb(i, j).y, 4)
        Call backbuffer.SetForeColor(oldcol)
        Call backbuffer.SetFillStyle(1)
    End If
Else
    If bomb(i, j).type = 9 Then
        rr.Left = 0
        rr.Top = 8 * maph
        rr.Right = mapw * 2
        rr.Bottom = 8 * maph + maph * 2
    ElseIf bomb(i, j).type = 10 Then
        rr.Left = 0
        rr.Top = 9 * maph
        rr.Right = mapw * 2
        rr.Bottom = 9 * maph + maph * 2
    End If
    If bomb(i, j).x < mapw Then
        bomb(i, j).x = mapw
        rr.Left = bomb(i, j).x
    End If
    If bomb(i, j).y < maph Then
        bomb(i, j).y = maph
        rr.Top = rr.Top + maph
    End If
    If bomb(i, j).y > 480 - maph Then
        rr.Bottom = rr.Bottom - (bomb(i, j).y - (480 - maph))
    End If
    Call backbuffer.BltFast(bomb(i, j).x - mapw, bomb(i, j).y - maph, mapsurf, rr, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
End If
End Sub
Sub killmap(tx, ty)
If Form1.WindowState = 1 Then Exit Sub
Call blacksurf.SetFillColor(vbBlack)
If map(tx, ty).halfx = 1 Then
    Call blacksurf.drawbox(map(tx, ty).x, map(tx, ty).y, map(tx, ty).x + mapw / 2, map(tx, ty).y + maph)
End If
If map(tx, ty).halfx = 2 Then
    Call blacksurf.drawbox(map(tx, ty).x + mapw / 2, map(tx, ty).y, map(tx, ty).x + mapw, map(tx, ty).y + maph)
End If
If map(tx, ty).halfx = 3 Or map(tx, ty).halfy = 3 Then
    Call blacksurf.drawbox(map(tx, ty).x, map(tx, ty).y, map(tx, ty).x + mapw, map(tx, ty).y + maph)
End If
If map(tx, ty).halfy = 1 Then
    Call blacksurf.drawbox(map(tx, ty).x, map(tx, ty).y, map(tx, ty).x + mapw, map(tx, ty).y + maph / 2)
End If
If map(tx, ty).halfy = 2 Then
    Call blacksurf.drawbox(map(tx, ty).x, map(tx, ty).y + maph / 2, map(tx, ty).x + mapw, map(tx, ty).y + maph)
End If

End Sub
Sub cutpic() '#########################################################
Dim i As Long
Const Maxboxcount = 30
Static Boxcount As Long
Boxcount = Boxcount + 1
If Boxcount >= Maxboxcount Then Boxcount = 0
For i = 1 To 10
    If tank(i).fangxiang = 2 Then
        tank(i).r.Left = 0
        tank(i).r.Right = ddsd3.lWidth * (40 / 160)
    ElseIf tank(i).fangxiang = 4 Then
        tank(i).r.Left = ddsd3.lWidth * (41 / 160)
        tank(i).r.Right = ddsd3.lWidth * (80 / 160)
    ElseIf tank(i).fangxiang = 1 Then
        tank(i).r.Left = ddsd3.lWidth * (81 / 160)
        tank(i).r.Right = ddsd3.lWidth * (120 / 160)
    Else:
        tank(i).r.Left = ddsd3.lWidth * (121 / 160)
        tank(i).r.Right = ddsd3.lWidth
    End If
    If tank(i).box = False Or Boxcount <= Maxboxcount / 3 Then
        tank(i).r.Top = (ddsd3.lHeight / 6) * (tank(i).shape - 1) + 1
        tank(i).r.Bottom = (ddsd3.lHeight / 6) * (tank(i).shape) - 1
    Else
        tank(i).r.Top = (ddsd3.lHeight / 6) * (5) + 1
        tank(i).r.Bottom = (ddsd3.lHeight / 6) * (6) - 1
    End If
Next i
End Sub

Sub EndIT()
    Dim r As Long
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    'r = timeKillEvent(timerid)
    End
End Sub

Private Sub Form_Click()
    'EndIT
End Sub
        
Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
jianpandown (keycode)


End Sub
Private Sub Form_Keyup(keycode As Integer, Shift As Integer)
jianpanup (keycode)
End Sub

Private Sub Form_Load()

    Init

End Sub

Private Sub Form_Paint()

    blt

End Sub

Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = dd.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
    
End Function

