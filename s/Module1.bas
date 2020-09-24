Attribute VB_Name = "Module1"
'sorry my english is poor
Option Explicit
Dim kb(1 To 12) As Long
Dim oldkb(2) As Long
Dim newkb(2) As Long
Public bug(5) As String
Public Type tanktype
    x As Single
    y As Single
    shape As Long
    fangxiang As Long
    star As Long
    gun As Long
    r As RECT
    wudi As Long
    linex As Long 'on the line
    liney As Long
    linex2 As Long
    liney2 As Long
    life As Long
    v As Long
    box As Boolean
End Type
Public bomb(1 To 10, 1 To 3) As bombtype

Type bombtype
    x As Long
    y As Long
    v As Long
    f As Long
    linex As Long
    liney As Long
    type As Long
End Type
Public tank(1 To 10) As tanktype
Type maptype
    x As Single
    y As Single
    shape As Long
    halfx As Long 'halfx=1 Ê±£¬µØÍ¼È±×ó±ß£¬=2Ê±£¬µØÍ¼È±ÓÒ±ß
    halfy As Long 'halfy=1 Ê±£¬µØÍ¼È±ÉÏ±ß£¬=2 Ê±£¬µØÍ¼È±ÏÂ±ß
End Type
Type Boxtype
    x As Single
    y As Single
    shape As Long
    timer As Long
End Type
Dim Notmovetimer As Long
Public Shuaiwudi As Long
Public box As Boxtype
Public mapw As Single, maph As Single
Public map(1 To 26, 1 To 23) As maptype
Const step = 2
Const bombstep = 4
Const tanklife = 5
Dim hadchushihua As Boolean
Dim player As Long
Dim shuai As Long
Dim gameisover(1 To 2) As Boolean
Dim enemycount As Long
Public enemyleft As Long
Dim enemyon As Long
Public score(1 To 2) As Long
Dim oldscore(1 To 2) As Long
Public round As Long
Dim ccfang(3 To 10) As Long 'Ëæ»úÊý£¬ÒÔ¾ö¶¨ÊÇ·ñ¸Ä±ä·½Ïò
Dim ccshoot(3 To 10) As Long 'Ëæ»úÊý£¬ÒÔ¾ö¶¨ÊÇ·ñÉä»÷
Public start As Long
Public passav As Integer
Public thanks(1 To 5) As String
Sub fight()
Static oldtime As Long
Static anotherbomb(10) As Long
Static kaishi As Long
Dim i As Long
Dim l As Long
If start = 0 Then Exit Sub
If passav <= 100 Then
    Form1.blt
    Exit Sub
End If
If thanks(1) <> "" Then
    Form1.blt
    Exit Sub
End If
If kaishi = 0 Then
    kaishi = 1
    enemycount = 20
    Randomize
    hadchushihua = False
    round = 1
    pass (round)
End If

If hadchushihua = False Then Exit Sub
If gameisover(1) = False Then
    If tank(1).shape <> 0 Then
        If kb(2) = 1 Then
            tank(1).y = tank(1).y + tank(1).v
            Call Turn(1, 1)
            tank(1).fangxiang = 1
            moveornot (1)
        ElseIf kb(1) = 1 Then
            tank(1).y = tank(1).y - tank(1).v
            Call Turn(1, 2)
            tank(1).fangxiang = 2
            moveornot (1)
        ElseIf kb(3) = 1 Then
            tank(1).x = tank(1).x - tank(1).v
            Call Turn(1, 3)
            tank(1).fangxiang = 3
            moveornot (1)
        ElseIf kb(4) = 1 Then
            tank(1).x = tank(1).x + tank(1).v
             Call Turn(1, 4)
             tank(1).fangxiang = 4
             moveornot (1)
        End If
        If kb(5) = 1 Then
            If bomb(1, 1).type <> bomb(1, 2).type Then anotherbomb(1) = anotherbomb(1) + 1
            If bomb(1, 1).type = bomb(1, 2).type Then anotherbomb(1) = 11
            If anotherbomb(1) > 10 Then
                Shoot (1)
                anotherbomb(1) = 0
            End If
        End If
    End If
End If
If gameisover(2) = False Then
    If tank(2).shape <> 0 Then
        If kb(8) = 1 Then
            tank(2).y = tank(2).y + tank(2).v
            
            Call Turn(2, 1)
            tank(2).fangxiang = 1
            moveornot (2)
        ElseIf kb(7) = 1 Then
            tank(2).y = tank(2).y - tank(2).v
            
            Call Turn(2, 2)
            tank(2).fangxiang = 2
            moveornot (2)
        ElseIf kb(9) = 1 Then
            tank(2).x = tank(2).x - tank(2).v
    
            Call Turn(2, 3)
            tank(2).fangxiang = 3
            moveornot (2)
        ElseIf kb(10) = 1 Then
            tank(2).x = tank(2).x + tank(2).v
            Call Turn(2, 4)
            tank(2).fangxiang = 4
            moveornot (2)
        End If
        If kb(11) = 1 Then
            If bomb(2, 1).type <> bomb(2, 2).type Then anotherbomb(2) = anotherbomb(2) + 1
            If bomb(2, 1).type = bomb(2, 2).type Then anotherbomb(2) = 11
    
            If anotherbomb(2) > 10 Then 'the time between two bomb
                Shoot (2)
                anotherbomb(2) = 0
            End If
        End If
    End If
End If
    Controlenemy
    Call Bombfly
    Form1.cutpic
    Form1.blt


End Sub
Sub Controlenemy() 'enemy
Dim i As Long
Static wait As Long
Static waitshow As Long
Dim count As Long
If enemyon < 4 * player And enemyleft > 0 Then
    Showenemy

End If

For i = 3 To 10
    If Notmovetimer > 0 Then
        Notmovetimer = Notmovetimer - 1
    Else
        If tank(i).shape <> 0 Then
            Call enemyturnandshoot(i)
            Call Trymove(i)
        End If
    End If
Next i
If enemyleft = 0 And enemyon = 0 Then
    wait = wait + 1
    
End If
If wait >= 200 Then
    wait = 0
    round = round + 1
    oldscore(1) = score(1)
    oldscore(2) = score(2)
    pass (round)
End If
End Sub
Sub enemyturnandshoot(i As Long)
    If ccfang(i) <= 0 Then ccfang(i) = Rnd * 100
    If ccfang(i) > 0 Then ccfang(i) = ccfang(i) - 1
    If ccshoot(i) <= 0 Then ccshoot(i) = Rnd * 80
    If ccshoot(i) > 0 Then ccshoot(i) = ccshoot(i) - 1
Dim fang As Long

    If tank(i).shape <> 0 Then
        If ccshoot(i) < 1 Then Call Shoot(i)
        fang = (Rnd * 6 Mod 6) + 1
        If fang >= 5 Then fang = 1
        If tank(i).liney >= 22 Then
            fang = (Rnd * 6 Mod 6) + 1
            If fang = 5 Then fang = 3
            If fang = 6 Then fang = 4
        End If
        If ccfang(i) <= 1 Then
            Call Turn(i, fang)
            tank(i).fangxiang = fang
           
        End If
        
    End If



End Sub
Sub Trymove(i As Long)
If tank(i).fangxiang = 1 Then tank(i).y = tank(i).y + tank(i).v
If tank(i).fangxiang = 2 Then tank(i).y = tank(i).y - tank(i).v
If tank(i).fangxiang = 3 Then tank(i).x = tank(i).x - tank(i).v
If tank(i).fangxiang = 4 Then tank(i).x = tank(i).x + tank(i).v
Call Turn(i, tank(i).fangxiang)
Call Turn2(i)
Call moveornot(i)

End Sub
Sub Showenemy()
Static chaotimer(1 To 3) As Long
Dim j As Long
Dim i As Long
Dim tx As Long
Dim shape As Long
For i = 1 To 3
    If chaotimer(i) > 0 Then
        chaotimer(i) = chaotimer(i) - 1
    End If
Next i
Dim chao As Long

For i = 3 To 4 * player + 2
    If tank(i).shape = 0 Then
        chao = ((Rnd * 3) Mod 3) + 1
        shape = (Rnd * 3 Mod 3) + 3
        If chaotimer(chao) = 0 Then
            chaotimer(chao) = 50
            If chao = 1 Then tx = 1
            If chao = 2 Then tx = 13
            If chao = 3 Then tx = 25
            enemyleft = enemyleft - 1
            enemyon = enemyon + 1
            Call Newtank(i, tx, 1, shape)
            Exit For
        End If
    End If
Next i
End Sub

Sub Bombfly() '×Óµ¯µÄÖ÷º¯Êý
Dim i As Long
Dim j As Long
For i = 1 To 10
    For j = 1 To 3
        If bomb(i, j).type = 9 Then
            bomb(i, j).type = 10
        ElseIf bomb(i, j).type = 10 Then
            bomb(i, j).type = 0
        End If

        If bomb(i, j).type > 0 And bomb(i, j).type < 9 Then
            If bomb(i, j).f = 2 Then
                bomb(i, j).y = bomb(i, j).y - bomb(i, j).v
            ElseIf bomb(i, j).f = 1 Then
                bomb(i, j).y = bomb(i, j).y + bomb(i, j).v
            ElseIf bomb(i, j).f = 3 Then
                bomb(i, j).x = bomb(i, j).x - bomb(i, j).v
            ElseIf bomb(i, j).f = 4 Then
                bomb(i, j).x = bomb(i, j).x + bomb(i, j).v
            End If
            If bombBaozhaornot(i, j) = True Then
            'ÔÚ´Ë´¦¼ÓÈëÉùÒô
            End If
        End If
    Next j
Next i
End Sub
Function bombBaozhaornot(i As Long, j As Long) As Boolean
Dim tx As Long
Dim ty As Long
Dim k As Long, m As Long
Dim hadkill As Boolean
hadkill = False
bombBaozhaornot = False
bomb(i, j).liney = (bomb(i, j).y + mapw + 5) \ (maph)
bomb(i, j).linex = (bomb(i, j).x + maph + 7) \ (mapw)

tx = bomb(i, j).linex
ty = bomb(i, j).liney
If tx > 26 Then
    bomb(i, j).linex = bomb(i, j).linex - 1
    bomb(i, j).type = 9
ElseIf tx < 1 Then
    bomb(i, j).linex = bomb(i, j).linex + 1
    bomb(i, j).type = 9
ElseIf ty > 23 Then
    bomb(i, j).liney = bomb(i, j).liney - 1
    bomb(i, j).type = 9
ElseIf ty < 1 Then
    bomb(i, j).liney = bomb(i, j).liney + 1
    bomb(i, j).type = 9
End If
If bomb(i, j).type = 9 Then Exit Function
'###########################################

If bomb(i, j).f = 1 Or bomb(i, j).f = 2 Then
    If (map(tx, ty).shape = 1 Or map(tx, ty).shape = 2) Then
        If map(tx, ty).halfx = 0 Or map(tx, ty).halfx = 2 Then
            bomb(i, j).type = 9
            If tank(i).gun = 1 Or map(tx, ty).shape = 1 Then
                If map(tx, ty).halfy = 0 Then
                    map(tx, ty).halfy = bomb(i, j).f
                ElseIf (map(tx, ty).halfy = 1 Or map(tx, ty).halfy = 2) Then ' And (map(tx - 1, ty).halfy = bomb(i, j).f Or map(tx - 1, ty).halfy = 3) Then
                    map(tx, ty).halfy = 3
                    map(tx, ty).shape = 0
                End If
                Call Form1.killmap(tx, ty)
            End If
        End If
    End If
    If (map(tx - 1, ty).shape = 1 Or map(tx - 1, ty).shape = 2) Then
        If map(tx - 1, ty).halfx = 0 Or map(tx - 1, ty).halfx = 1 Then
            bomb(i, j).type = 9
            If tank(i).gun = 1 Or map(tx - 1, ty).shape = 1 Then
                If map(tx - 1, ty).halfy = 0 Then
                    map(tx - 1, ty).halfy = bomb(i, j).f
                ElseIf (map(tx - 1, ty).halfy = 1 Or map(tx - 1, ty).halfy = 2) Then ' And map(tx, ty).halfy <> bomb(i, j).f Then
                    map(tx - 1, ty).halfy = 3
                    map(tx - 1, ty).shape = 0
                End If
                Call Form1.killmap(tx - 1, ty)
            End If
        End If
    End If
ElseIf bomb(i, j).f = 3 Or bomb(i, j).f = 4 Then
    If (map(tx, ty).shape = 1 Or map(tx, ty).shape = 2) Then
        If map(tx, ty).halfy = 0 Or map(tx, ty).halfy = 2 Then
            bomb(i, j).type = 9
            If tank(i).gun = 1 Or map(tx, ty).shape = 1 Then
                If map(tx, ty).halfx = 0 Then
                    map(tx, ty).halfx = 5 - bomb(i, j).f
                ElseIf (map(tx, ty).halfx = 1 Or map(tx, ty).halfx = 2) Then 'And (map(tx, ty - 1).halfx = 5 - bomb(i, j).f Or map(tx, ty - 1).halfx = 3) Then
                    map(tx, ty).halfx = 3
                    map(tx, ty).shape = 0
                End If
                Call Form1.killmap(tx, ty)
            End If
        End If
    End If
    If (map(tx, ty - 1).shape = 1 Or map(tx, ty - 1).shape = 2) Then
        If map(tx, ty - 1).halfy = 0 Or map(tx, ty - 1).halfy = 1 Then
            bomb(i, j).type = 9
            If tank(i).gun = 1 Or map(tx, ty - 1).shape = 1 Then
                If map(tx, ty - 1).halfx = 0 Then
                    map(tx, ty - 1).halfx = 5 - bomb(i, j).f
                ElseIf (map(tx, ty - 1).halfx = 1 Or map(tx, ty - 1).halfx = 2) Then ' And map(tx, ty).halfx <> 5 - bomb(i, j).f Then
                    map(tx, ty - 1).shape = 0
                    map(tx, ty - 1).halfy = 3
                End If
                Call Form1.killmap(tx, ty - 1)
            End If
        End If
    End If
End If
'##########################'check the base
For k = 22 To 23
    For m = 13 To 14
        If map(m, k).halfx <> 0 Or map(m, k).halfy <> 0 Then
            gameover
        End If
    Next m
Next k

'##########################'tank was kill
For k = 1 To 10
    If tank(k).shape <> 0 Then
        If (i = 1 And k <> 2 And k <> 1) Or (i = 2 And k <> 1 And k <> 2) Or (i >= 3 And k <= 2 And k <> i) Or (k >= 3 And i <= 2 And k <> i) Then
       'If (i = 1 And k <> 1) Then
            If bomb(i, j).linex = bomb(k, 1).linex And bomb(i, j).liney = bomb(k, 1).liney And bomb(k, 1).type <> 0 Then
                If tank(i).gun = 0 Then bomb(i, j).type = 0
                bomb(k, 1).type = 0
            End If
            If bomb(i, j).linex = bomb(k, 2).linex And bomb(i, j).liney = bomb(k, 2).liney And bomb(k, 2).type < 0 Then
                If tank(i).gun = 0 Then bomb(i, j).type = 0
                bomb(k, 2).type = 0
            End If
            If bomb(i, j).f = 1 Or bomb(i, j).f = 2 Then

                If bomb(i, j).x - tank(k).x >= -mapw - 5 And bomb(i, j).x - tank(k).x <= mapw + 5 Then
                    If Abs(bomb(i, j).y - tank(k).y) <= maph Then
                        'bomb(i, j).y = tank(k).y
                        bomb(k, 3).x = tank(k).x
                        bomb(k, 3).y = tank(k).y
                        bomb(k, 3).type = 9
                        bomb(i, j).type = 9
                        Call tankdied(k, i)
                    End If
                End If
            ElseIf bomb(i, j).f = 3 Or bomb(i, j).f = 4 Then
                If bomb(i, j).y - tank(k).y >= -maph - 5 And bomb(i, j).y - tank(k).y <= maph + 5 Then
                    If Abs(bomb(i, j).x - tank(k).x) <= mapw Then
                        'bomb(i, j).x = tank(k).x
                        bomb(i, j).type = 9
                        Call tankdied(k, i)
                    End If
                End If
            End If
        End If
    End If
Next k

    
End Function
Sub gameover()


    bug(1) = "game is over"
    bug(2) = "press esc to exit , F1 or F2 to restart ,F3 to continue"
    If gameisover(1) + gameisover(2) <> 0 Then Call Form1.drawshuai(0)
    gameisover(1) = True
    gameisover(2) = True


End Sub
Sub tankdied(index As Long, killer As Long)
If tank(index).wudi > 0 Then Exit Sub
If tank(index).shape = 0 Then Exit Sub
If killer = 0 Then
    If index >= 3 Then
        bomb(index, 3).x = tank(index).x
        bomb(index, 3).y = tank(index).y
        bomb(index, 3).type = 9
        tank(index).shape = 0
        enemyon = enemyon - 1
    End If
Else
    If index >= 3 Then
        If tank(index).shape = 3 Or tank(killer).gun = 1 Then
            tank(index).shape = 0
            enemyon = enemyon - 1
            score(killer) = score(killer) + 100
        End If
        If tank(index).shape > 3 And tank(killer).gun <> 1 Then
            tank(index).shape = tank(index).shape - 1
            score(killer) = score(killer) + 100
            If score(killer) <> 0 And score(killer) Mod 10000 = 0 Then tank(killer).life = tank(killer).life + 1
        End If
        If tank(index).box = True Then
            Showbox
        End If
    End If
End If
If index = 1 Then
    If tank(index).gun = 0 And tank(index).star = 0 Then
        tank(1).shape = 0
        tank(1).life = tank(1).life - 1
        Call Newtank(index, 1, 1, 1)
    Else
        tank(1).gun = 0
        tank(1).star = 0
    End If
    kb(5) = 0
    
End If

If index = 2 And player = 2 Then
    If tank(2).gun = 0 And tank(2).star = 0 Then
        tank(2).shape = 0
        tank(2).life = tank(2).life - 1
        Call Newtank(index, 1, 1, 1)
    Else
        tank(2).gun = 0
        tank(2).star = 0
    End If
    kb(11) = 0
    
End If
End Sub
Sub Showbox()
Dim rndy As Long
Dim rndx As Long
rndx = (Rnd * 25) Mod 25 + 1
rndy = (Rnd * 21) Mod 21 + 1
box.x = map(rndx, rndy).x
box.y = map(rndx, rndy).y
box.shape = (Rnd * 7 Mod 7) + 1
box.timer = 500
End Sub

Sub Newtank(index, linex, liney, shape)
If index >= 3 Then
    tank(index).fangxiang = 1
    tank(index).gun = 0
    tank(index).linex = linex
    tank(index).liney = liney
    tank(index).star = 0
    tank(index).wudi = 0
    tank(index).x = map(tank(index).linex + 1, 10).x
    tank(index).y = map(10, tank(index).liney + 1).y + 3
    tank(index).linex2 = (tank(index).x + mapw + 3) \ (mapw)
    tank(index).liney2 = (tank(index).y + maph + 3) \ (maph)
    tank(index).v = step
    tank(index).shape = shape
    tank(index).box = False
    If Rnd * 20 < 5 Then tank(index).box = True
Else
    If tank(1).life <= 0 And tank(2).life <= 0 Then gameover
    If tank(index).life >= 1 Then
        
        If index = 1 Then
            tank(index).linex = 10
            tank(index).shape = 1
        Else
            tank(index).linex = 17
            tank(index).shape = 2
        End If
        tank(index).fangxiang = 2
        tank(index).gun = 0
        tank(index).liney = 22
        tank(index).star = 0
        tank(index).wudi = 200
        tank(index).x = map(tank(index).linex + 1, 10).x
        tank(index).y = map(10, tank(index).liney + 1).y + 3
        tank(index).linex2 = (tank(index).x + mapw + 3) \ (mapw)
        tank(index).liney2 = (tank(index).y + maph + 3) \ (maph)
        tank(index).v = step
    End If
End If

End Sub
Sub Borrowlife(index As Long)
If tank(3 - index).life >= 2 Then
    tank(3 - index).life = tank(3 - index).life - 1
    tank(index).life = tank(index).life + 1
    Call Newtank(index, 1, 1, 1)
End If
End Sub
Sub Shoot(index As Long)
If tank(index).shape = 0 Then
    If index <= 2 And player = 2 And tank(index).life = 0 Then Call Borrowlife(index)
    Exit Sub
End If
If index <= 2 Then

    If gameisover(index) = True Then Exit Sub
End If

If bomb(index, 1).type = 0 Then
    If tank(index).gun = 0 Then
        bomb(index, 1).type = 1
    ElseIf tank(index).gun = 1 Then
        bomb(index, 1).type = 1
    End If
    If tank(index).star = 1 Then bomb(index, 1).v = tank(index).v + 2
    bomb(index, 1).f = tank(index).fangxiang
    If tank(index).fangxiang = 1 Then
        bomb(index, 1).x = tank(index).x
        bomb(index, 1).linex = tank(index).linex
        bomb(index, 1).y = tank(index).y + maph - 10
    ElseIf tank(index).fangxiang = 2 Then
        bomb(index, 1).x = tank(index).x
        bomb(index, 1).linex = tank(index).linex
        bomb(index, 1).y = tank(index).y - maph + 10
    ElseIf tank(index).fangxiang = 3 Then
        bomb(index, 1).x = tank(index).x - mapw + 10
        bomb(index, 1).y = tank(index).y
        bomb(index, 1).liney = tank(index).liney
    ElseIf tank(index).fangxiang = 4 Then
        bomb(index, 1).x = tank(index).x + mapw - 10
        bomb(index, 1).y = tank(index).y
        bomb(index, 1).liney = tank(index).liney
    End If
ElseIf bomb(index, 2).type = 0 And tank(index).star = 1 Then
        If tank(index).gun = 0 Then
            bomb(index, 2).type = 1
        ElseIf tank(index).gun = 1 Then
            bomb(index, 2).type = 1
        End If
        bomb(index, 2).v = tank(index).v + 2
    bomb(index, 2).f = tank(index).fangxiang
    If tank(index).fangxiang = 1 Then
        bomb(index, 2).x = tank(index).x
        bomb(index, 2).y = tank(index).y + mapw
        bomb(index, 2).linex = tank(index).linex
    ElseIf tank(index).fangxiang = 2 Then
        bomb(index, 2).x = tank(index).x
        bomb(index, 2).linex = tank(index).linex
        bomb(index, 2).y = tank(index).y - mapw
    ElseIf tank(index).fangxiang = 3 Then
        bomb(index, 2).x = tank(index).x - mapw
        bomb(index, 2).y = tank(index).y
        bomb(index, 2).liney = tank(index).liney
    ElseIf tank(index).fangxiang = 4 Then
        bomb(index, 2).x = tank(index).x + mapw
        bomb(index, 2).y = tank(index).y
        bomb(index, 2).liney = tank(index).liney
    End If
    
End If

End Sub
Sub moveornot(index As Long)
Dim i As Long, k As Long
If tank(index).shape = 0 Then Exit Sub
i = index
Turn2 (i)
If notmove(i) = 1 Then tank(i).y = tank(i).y + tank(i).v
If tank(i).y = 480 - mapw - 5 Then tank(i).y = 480 - mapw - 5
If notmove(i) = 2 Then tank(i).y = tank(i).y - tank(i).v
If tank(i).y < mapw + 5 Then tank(i).y = mapw + 5
If notmove(i) = 3 Then tank(i).x = tank(i).x + tank(i).v
If tank(i).x = 520 - mapw - 5 Then tank(i).x = 520 - mapw - 5
If notmove(i) = 4 Then tank(i).x = tank(i).x - tank(i).v
If tank(i).x < mapw + 5 Then tank(i).x = mapw + 5
If i >= 3 And notmove(i) <> 0 Then
    ccfang(i) = ccfang(i) - 2
End If
If i <= 2 And box.shape <> 0 Then '¼ñºÐ×Ó
    If Abs(tank(i).x - box.x) <= 2 * mapw - 10 And Abs(tank(i).y - box.y) <= 2 * maph - 10 Then
        score(i) = score(i) + 500
        If box.shape = 1 Then tank(i).wudi = 800
        If box.shape = 2 Then
            For k = 3 To 10
                If tank(k).shape <> 0 Then
                    Call tankdied(k, 0)
                    
                End If
            Next k
        End If
        If box.shape = 3 Then
            Shuaiwudi = 1500
        End If
        If box.shape = 4 Then tank(i).life = tank(i).life + 1
        If box.shape = 5 Then
            If tank(i).star = 1 Then
                tank(i).gun = 1
            Else
                tank(i).star = 1
            End If
            tank(i).v = step + 1
            
                
        End If
        If box.shape = 6 Then
            Notmovetimer = 4000
        End If
        If box.shape = 7 Then
            tank(i).gun = 1
            tank(i).star = 1
            tank(i).v = step + 1
        End If
        box.shape = 0
        box.timer = 0
    End If
    
End If
End Sub

Sub Turn2(index As Long)
If tank(index).shape = 0 Then Exit Sub
    tank(index).linex2 = (tank(index).x + mapw + 3) \ (mapw)
    tank(index).liney2 = (tank(index).y + maph + 3) \ (maph)

    
End Sub
Sub Turn(index As Long, fangxiang As Long)
If tank(index).shape = 0 Then Exit Sub
tank(index).linex = (tank(index).x + mapw / 2) \ (mapw)
tank(index).liney = (tank(index).y + maph / 2) \ (maph)
    
If fangxiang <> tank(index).fangxiang Then
    tank(index).x = map(tank(index).linex + 1, 10).x
    tank(index).y = map(10, tank(index).liney + 1).y + 3
    'tank(index).fangxiang = fangxiang
End If

End Sub

Function notmove(index As Long) As Long
Dim tx As Long, ty As Long
Dim i As Long, j As Long
If tank(index).shape = 0 Then Exit Function
tx = tank(index).linex2
ty = tank(index).liney2
notmove = 0
If tank(index).y <= maph + 1 Then
    notmove = 1
ElseIf tank(index).y >= 487 - maph - 1 Then
    notmove = 2
ElseIf tank(index).x <= mapw + 1 Then
    notmove = 3
ElseIf tank(index).x >= 525 - mapw - 1 Then
    notmove = 4
End If
For i = 1 To 2 '1 ºÍ 2 ÎªÌ¹¿Ë×óÇ°·½ºÍÓÒÇ°·½
    Select Case tank(index).fangxiang
        Case 2
            If ty > 1 Then
                If map(tx + 1 - i, ty - 1).shape <> 0 And map(tx + 1 - i, ty - 1).shape <= 3 Then
                    notmove = 1
                End If
                For j = 1 To 10
                    If j <> index And i = 1 And Abs(tank(index).linex2 - tank(j).linex2) <= 1 Then
                        If tank(index).y - tank(j).y < maph * 2 And tank(index).y - tank(j).y > mapw And tank(j).shape <> 0 Then notmove = 1
                        
                    End If
                Next j
            End If
        Case 1
            If ty < 23 Then
                If map(tx + 1 - i, ty + 1).shape <> 0 And map(tx + 1 - i, ty + 1).shape <= 3 Then
                    notmove = 2
                End If
                For j = 1 To 10
                    If j <> index And i = 1 And Abs(tank(index).linex2 - tank(j).linex2) <= 1 Then
                        If tank(j).y - tank(index).y < maph * 2 And tank(j).y - tank(index).y > mapw And tank(j).shape <> 0 Then notmove = 2
                    End If
                Next j
            End If
        Case 3
            If tx > 1 Then
                If map(tx - 1, ty + 1 - i).shape <> 0 And map(tx - 1, ty + 1 - i).shape <= 3 Then
                    notmove = 3
                End If
                For j = 1 To 10
                    If j <> index And i = 1 And Abs(tank(index).liney2 - tank(j).liney2) <= 1 Then
                        If tank(index).x - tank(j).x < maph * 2 And tank(index).x - tank(j).x > mapw And tank(j).shape <> 0 Then notmove = 3
                    End If
                Next j

            End If
        Case 4
            If tx < 26 Then
                If map(tx + 1, ty + 1 - i).shape <> 0 And map(tx + 1, ty + 1 - i).shape <= 3 Then
                    notmove = 4
                End If
                For j = 1 To 10
                    If j <> index And i = 1 And Abs(tank(index).liney2 - tank(j).liney2) <= 1 Then
                        If tank(j).x - tank(index).x < maph * 2 And tank(j).x - tank(index).x > mapw And tank(j).shape <> 0 Then notmove = 4
                    End If
                Next j
                
            End If
    End Select
Next i
End Function
Sub Thanksplayer()
    thanks(1) = "you had passed all"
    thanks(2) = "this is my vb game"
    thanks(3) = "hope to make frind to you                    wang hailong"
    thanks(4) = "                                  callwachel@hotmail.com"
    gameover
End Sub
Sub pass(round As Long)
Dim mapfilename As String
Dim s As String
Dim i As Long, j As Long, k As Long

mapfilename = "map" & "\" & "map" & round & ".txt"
If Dir(mapfilename) = "" Then
    round = 0
    Thanksplayer
    
    Exit Sub
End If
Open mapfilename For Input As #1
    Input #1, s
Close #1
mapw = 520 / 26
maph = 480 / 23
For j = 1 To 23
    For i = 1 To 26
        map(i, j).x = (i - 1) * mapw
        map(i, j).y = (j - 1) * maph
    Next i
Next j
For j = 1 To 23
    For i = 1 To 26
        map(i, j).halfx = 3
        map(i, j).halfy = 3
        map(i, j).shape = Mid(s, i + (j - 1) * 26, 1)
        If map(i, j).shape = 1 Or map(i, j).shape = 2 Then
            map(i, j).halfx = 0
            map(i, j).halfy = 0
        End If
        Call Form1.drawmap(map(i, j).x, map(i, j).y, map(i, j).shape)
    Next i
Next j
Chushihua
End Sub
Sub jianpandown(keycode As Long)
Select Case keycode
    Case vbKeyW
        kb(1) = 1
    Case vbKeyUp
        kb(7) = 1
    Case vbKeyS
        kb(2) = 1
    Case vbKeyDown
        kb(8) = 1
    Case vbKeyA
        kb(3) = 1
    Case vbKeyLeft
        kb(9) = 1
    Case vbKeyD
        kb(4) = 1
    Case vbKeyRight
        kb(10) = 1
    Case vbKeyJ
        kb(5) = 1
        Call Shoot(1)
    Case vbKeyNumpad0
        kb(11) = 1
        Call Shoot(2)
    Case vbKeyK
        kb(6) = 1
    Case vbKeyDelete
        kb(12) = 1
End Select
fight
If keycode = vbKeyEscape Then Call Form1.EndIT
If start = 1 Then
    If keycode = vbKeyF1 Then
        hadchushihua = False
        player = 1
        passav = 0
        round = 1
        pass (round)
        tank(2).shape = 0
        tank(2).life = 0
        tank(1).life = tanklife
        tankchushihua (1)
    End If
    If keycode = vbKeyF2 Then
        hadchushihua = False
        player = 2
        round = 1
        passav = 0
        pass (round)
        tank(1).life = tanklife
        tank(2).life = tanklife
        tankchushihua (1)
        tankchushihua (2)
    End If
    If keycode = vbKeyF3 Then
        score(1) = oldscore(1)
        score(2) = oldscore(2)
        hadchushihua = False
        passav = 0
        pass (round)
        tank(1).life = tanklife
        tankchushihua (1)
        If player = 2 Then
            tank(2).life = tanklife
            tankchushihua (2)
        End If
    End If
End If
If start = 0 Then
    If keycode = vbKeyReturn Then
        player = 1
        Form1.Drawpassav
        start = 1
    End If
    If keycode = vbKeyF1 Then
        player = 1
        Form1.Drawpassav
        start = 1
    End If
    If keycode = vbKeyF2 Then
        player = 2
        Form1.Drawpassav
        start = 1
    End If

End If
End Sub
Sub jianpanup(keycode As Long)
Select Case keycode
    Case vbKeyW
        kb(1) = 0
    Case vbKeyUp
        kb(7) = 0
    Case vbKeyS
        kb(2) = 0
    Case vbKeyDown
        kb(8) = 0
    Case vbKeyA
        kb(3) = 0
    Case vbKeyLeft
        kb(9) = 0
    Case vbKeyD
        kb(4) = 0
    Case vbKeyRight
        kb(10) = 0
    Case vbKeyJ
        kb(5) = 0
    Case vbKeyNumpad0
        kb(11) = 0
    Case vbKeyK
        kb(6) = 0
    Case vbKeyDelete
        kb(12) = 0
End Select
End Sub

Sub Chushihua()
Dim i As Long, j As Long
gameisover(1) = False
gameisover(2) = False
shuai = 1
box.shape = 0
Call Form1.drawshuai(1)
For j = 22 To 23
    For i = 13 To 14
        map(i, j).shape = 1
        map(i, j).halfx = 0
        map(i, j).halfy = 0
    Next i
Next j

For i = 3 To 10
    tank(i).shape = 0
    tank(i).star = 0
    tank(i).gun = 0
    tank(i).wudi = 0
    bomb(i, 1).v = bombstep
    bomb(i, 2).v = bombstep
    tank(i).x = map(11, 1).x
    tank(i).y = map(1, 23).y
    tank(i).fangxiang = 1
    tank(i).fangxiang = 1
    tank(i).wudi = 0
    tank(i).wudi = 0
    tank(i).star = 0
    tank(i).star = 0
Next i

For i = 1 To 10
    For j = 1 To 2
        bomb(i, j).type = 0
    Next j
Next i
'########
tank(1).x = map(11, 1).x
tank(1).y = map(1, 23).y
tank(1).fangxiang = 2
If player = 2 Then
    tank(2).x = map(17, 1).x
    tank(2).y = map(1, 23).y
    tank(2).fangxiang = 2
End If
'########
enemyleft = enemycount
enemyon = 0
Shuaiwudi = 0
bug(1) = ""
bug(2) = ""
If round = 1 Then
    score(1) = 0
    score(2) = 0
End If
thanks(1) = ""
hadchushihua = True
If round = 1 Then
    tankchushihua (1)
    If player = 2 Then tankchushihua (2)
    tank(1).life = tanklife
    If player = 2 Then tank(2).life = tanklife
End If
If round <> 1 Then passav = 0

End Sub
Sub tankchushihua(index As Long)
If index = 1 Then
    tank(1).shape = 1
    tank(1).star = 0
    tank(1).gun = 0
    tank(1).wudi = 200
    bomb(1, 1).v = bombstep
    bomb(1, 2).v = bombstep

    tank(1).star = 0
    tank(1).fangxiang = 3
    tank(1).v = step
End If
If index = 2 And player = 2 Then
    tank(2).shape = 2
    tank(2).star = 0
    bomb(2, 1).v = bombstep + 1
    bomb(2, 2).v = bombstep + 1
    '##
    tank(2).shape = 2
    tank(2).gun = 0

    tank(2).fangxiang = 3
    tank(2).wudi = 200
    tank(2).v = step
    tank(2).star = 0
End If
Call Turn(1, 4)
Call Turn(2, 4)
tank(1).fangxiang = 2
tank(2).fangxiang = 2
Call Turn(1, 1)
Call Turn(2, 1)
'##

End Sub
