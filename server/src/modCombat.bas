Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Integer, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(Index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (GetPlayerStat(Index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Integer, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal Index As Integer) As Long
    Dim weaponNum As Integer
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(Index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) / 5)
    End If

End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(npcNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (NPC(npcNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (NPC(npcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * NPC(npcNum).Stat(Stats.Strength) * NPC(npcNum).Damage + (NPC(npcNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(Index, Agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(Index, Agility) / 83.3
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal Index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(Index, Strength) * 0.25
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim Stat As Long
Dim rndNum As Long

    CanNpcBlock = False
    
    Stat = NPC(npcNum).Stat(Stats.Agility) / 5  'guessed shield agility
    rate = Stat / 12.08
    
    rndNum = rand(1, 100)
    
    If rndNum <= rate Then
        CanNpcBlock = True
    End If
    
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = NPC(npcNum).Stat(Stats.Agility) / 52.08
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = NPC(npcNum).Stat(Stats.Agility) / 83.3
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = NPC(npcNum).Stat(Stats.Strength) * 0.25
    rndNum = rand(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, mapNpcNum) Then
    
        mapnum = GetPlayerMap(Index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (NPC(npcNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, mapNpcNum, Damage)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(MapNpc(mapnum).NPC(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(mapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(mapNpcNum).y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackNpc = True
                    Else
                        ' ALATAR
                        If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            Call CheckTasks(attacker, QUEST_TYPE_GOTALK, npcNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, npcNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGET, npcNum)
                            
                            If NPC(npcNum).Quest = YES Then
                                If Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED Then
                                    If Quest(NPC(npcNum).Quest).Repeat = YES Then
                                        Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED_BUT
                                        Exit Function
                                    End If
                                End If
                                If CanStartQuest(attacker, NPC(npcNum).QuestNum) Then
                                    'if can start show the request message (speech1)
                                    QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).StartMessage), NPC(npcNum).QuestNum
                                    Exit Function
                                End If
                                If QuestInProgress(attacker, NPC(npcNum).QuestNum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).MidMessage), 0
                                    Exit Function
                                End If
                            End If
                        End If
                        
                        If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                            PlayerMsg attacker, Trim$(NPC(npcNum).Name) & ": " & Trim$(NPC(npcNum).AttackSay), White
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).Num
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        Exp = NPC(npcNum).Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, Exp, attacker, GetPlayerMap(attacker)
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, Exp
        End If
        
        'Drop the goods if they get it
        n = Int(Rnd * NPC(npcNum).DropChance) + 1

        If n = 1 Then
            Call SpawnItem(NPC(npcNum).DropItem, NPC(npcNum).DropItemValue, mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).Num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, False
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' ALATAR
        Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, npcNum)
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = mapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
        MapNpc(mapnum).NPC(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(i).Num = MapNpc(mapnum).NPC(mapNpcNum).Num Then
                    MapNpc(mapnum).NPC(i).target = attacker
                    MapNpc(mapnum).NPC(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
        MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellNum, attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    End If

    If spellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long)
Dim mapnum As Long, npcNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, Index) Then
        mapnum = GetPlayerMap(Index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(Index) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(Index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(Index, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, Index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).NPC(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Index)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).NPC(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapnum).NPC(mapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(Index) = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim mapnum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(mapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
    MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).Num
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        mapnum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellNum As Long = 0)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        ' send animation
        If n > 0 Then
            If spellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If

        If Exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - Exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & Exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, Exp, attacker, GetPlayerMap(attacker)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, Exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        ' ALATAR
        Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)

        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If spellNum = 0 Then Call SendAnimation(GetPlayerMap(victim), Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        End If
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Player victim, spellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellSlot As Long)
    Dim spellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellNum = GetPlayerSpell(Index, spellSlot)
    mapnum = GetPlayerMap(Index)
    
    If spellNum <= 0 Or spellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, spellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(spellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = Spell(spellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapnum).NPC(target).x, MapNpc(mapnum).NPC(target).y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg mapnum, "Casting " & Trim$(Spell(spellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        TempPlayer(Index).spellBuffer.Spell = spellSlot
        TempPlayer(Index).spellBuffer.Timer = GetTickCount
        TempPlayer(Index).spellBuffer.target = TempPlayer(Index).target
        TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer Index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellSlot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim spellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
   
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellSlot <= 0 Or spellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    spellNum = GetPlayerSpell(Index, spellSlot)
    mapnum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellNum) Then Exit Sub

    MPCost = Spell(spellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(spellNum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(spellNum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    Vital = Spell(spellNum).Vital
    AoE = Spell(spellNum).AoE
    Range = Spell(spellNum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(spellNum).Map, Spell(spellNum).x, Spell(spellNum).y
                    SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(mapnum).NPC(target).x
                    y = MapNpc(mapnum).NPC(target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    SendClearSpellBuffer Index
                End If
            End If
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(Index, i, True) Then
                                            SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer Index, i, Vital, spellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanPlayerAttackNpc(Index, i, True) Then
                                        SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc Index, i, Vital, spellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, spellNum
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellNum, mapnum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(mapnum).NPC(target).x
                y = MapNpc(mapnum).NPC(target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
           
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer Index, target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(Index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc Index, target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, spellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, spellNum
                        End If
                    Else
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(Index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, spellNum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, spellNum, mapnum
                        End If
                    End If
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
       
        TempPlayer(Index).SpellCD(spellSlot) = GetTickCount + (Spell(spellNum).CDTime * 1000)
        Call SendCooldown(Index, spellSlot)
        SendActionMsg mapnum, Trim$(Spell(spellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' send the sound
        SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, spellNum
        
        If increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Player Index, spellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal spellNum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapnum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg mapnum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(Index).x * 32, MapNpc(mapnum).NPC(Index).y * 32
        
        ' send the sound
        SendMapSound Index, MapNpc(mapnum).NPC(Index).x, MapNpc(mapnum).NPC(Index).y, SoundEntity.seSpell, spellNum
        
        If increment Then
            MapNpc(mapnum).NPC(Index).Vital(Vital) = MapNpc(mapnum).NPC(Index).Vital(Vital) + Damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Npc mapnum, Index, spellNum
            End If
        ElseIf Not increment Then
            MapNpc(mapnum).NPC(Index).Vital(Vital) = MapNpc(mapnum).NPC(Index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal spellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal Index As Long, ByVal spellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal spellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(Index).DoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal spellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(Index).HoT(i)
            If .Spell = spellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    PlayerAttackPlayer .Caster, Index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                   SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                   SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Spell(.Spell).Vital
                Else
                   SendActionMsg Player(Index).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(Index).x * 32, Player(Index).y * 32
                   SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal dotNum As Long)
    With MapNpc(mapnum).NPC(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, Index, True) Then
                    PlayerAttackNpc .Caster, Index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal Index As Long, ByVal hotNum As Long)
    With MapNpc(mapnum).NPC(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(Index).x * 32, MapNpc(mapnum).NPC(Index).y * 32
                    MapNpc(mapnum).NPC(Index).Vital(Vitals.HP) = MapNpc(mapnum).NPC(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                Else
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(Index).x * 32, MapNpc(mapnum).NPC(Index).y * 32
                    MapNpc(mapnum).NPC(Index).Vital(Vitals.MP) = MapNpc(mapnum).NPC(Index).Vital(Vitals.MP) + Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(spellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal mapnum As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(Index).StunDuration = Spell(spellNum).StunDuration
        MapNpc(mapnum).NPC(Index).StunTimer = GetTickCount
    End If
End Sub

