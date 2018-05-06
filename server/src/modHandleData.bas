Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPartyMsg) = GetAddress(AddressOf HandlePartyMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CSendPlusSkill) = GetAddress(AddressOf HandlePlusSkill)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    
    HandleDataSub(CRequestEditor) = GetAddress(AddressOf HandleRequestEditor)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(Player(Index).Name)) > 0 Then
                Call DeleteName(Player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(Index, Name)
            ClearBank Index
            LoadBank Index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            Else
                ' send new char shit
                If Not IsPlaying(Index) Then
                    Call SendNewCharClasses(Index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar Index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, White)
    Call SendChatBubble(GetPlayerMap(Index), Index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, Cyan)
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePartyMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    s = "[P]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Party(Index, Msg, BrightBlue)
    Call AddLog("[GRUPO]:" & Msg, GetPlayerName(Index) & ".log")
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim X As Long, Y As Long
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    'SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, X, Y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim Buffer As clsBuffer
Dim sMes As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(Index, PointType) >= 255 Then
            PlayerMsg Index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerRawStat(Index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(Index, Stats.Endurance, GetPlayerRawStat(Index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(Index, Stats.Intelligence, GetPlayerRawStat(Index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(Index, Stats.Agility, GetPlayerRawStat(Index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(Index, Stats.Willpower, GetPlayerRawStat(Index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(Index), "+1 " & sMes, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData Index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapNum As Long
    Dim X As Long
    Dim Y As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapNum = GetPlayerMap(Index)
    i = Map(mapNum).Revision + 1
    Call ClearMap(mapNum)
    
    Map(mapNum).Name = Buffer.ReadString
    Map(mapNum).Music = Buffer.ReadString
    Map(mapNum).BGS = Buffer.ReadString
    Map(mapNum).Revision = i
    Map(mapNum).Moral = Buffer.ReadByte
    Map(mapNum).Up = Buffer.ReadLong
    Map(mapNum).Down = Buffer.ReadLong
    Map(mapNum).Left = Buffer.ReadLong
    Map(mapNum).Right = Buffer.ReadLong
    Map(mapNum).BootMap = Buffer.ReadLong
    Map(mapNum).BootX = Buffer.ReadByte
    Map(mapNum).BootY = Buffer.ReadByte
    
    Map(mapNum).Weather = Buffer.ReadLong
    Map(mapNum).WeatherIntensity = Buffer.ReadLong
    
    Map(mapNum).Fog = Buffer.ReadLong
    Map(mapNum).FogSpeed = Buffer.ReadLong
    Map(mapNum).FogOpacity = Buffer.ReadLong
    
    Map(mapNum).Red = Buffer.ReadLong
    Map(mapNum).Green = Buffer.ReadLong
    Map(mapNum).Blue = Buffer.ReadLong
    Map(mapNum).Alpha = Buffer.ReadLong
    
    Map(mapNum).MaxX = Buffer.ReadByte
    Map(mapNum).MaxY = Buffer.ReadByte
    ReDim Map(mapNum).Tile(0 To Map(mapNum).MaxX, 0 To Map(mapNum).MaxY)

    For X = 0 To Map(mapNum).MaxX
        For Y = 0 To Map(mapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(mapNum).Tile(X, Y).Layer(i).X = Buffer.ReadLong
                Map(mapNum).Tile(X, Y).Layer(i).Y = Buffer.ReadLong
                Map(mapNum).Tile(X, Y).Layer(i).Tileset = Buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(mapNum).Tile(X, Y).Autotile(z) = Buffer.ReadLong
            Next
            Map(mapNum).Tile(X, Y).Type = Buffer.ReadByte
            Map(mapNum).Tile(X, Y).Data1 = Buffer.ReadLong
            Map(mapNum).Tile(X, Y).Data2 = Buffer.ReadLong
            Map(mapNum).Tile(X, Y).Data3 = Buffer.ReadLong
            Map(mapNum).Tile(X, Y).Data4 = Buffer.ReadString
            Map(mapNum).Tile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map(mapNum).NPC(X) = Buffer.ReadLong
        Map(mapNum).NpcSpawnType(X) = Buffer.ReadLong
        Call ClearMapNpc(X, mapNum)
    Next
    
    'Event Data!
    Map(mapNum).EventCount = Buffer.ReadLong
        
    If Map(mapNum).EventCount > 0 Then
        ReDim Map(mapNum).Events(0 To Map(mapNum).EventCount)
        For i = 1 To Map(mapNum).EventCount
            With Map(mapNum).Events(i)
                .Name = Buffer.ReadString
                .Global = Buffer.ReadLong
                .X = Buffer.ReadLong
                .Y = Buffer.ReadLong
                .PageCount = Buffer.ReadLong
            End With
            If Map(mapNum).Events(i).PageCount > 0 Then
                ReDim Map(mapNum).Events(i).Pages(0 To Map(mapNum).Events(i).PageCount)
                For X = 1 To Map(mapNum).Events(i).PageCount
                    With Map(mapNum).Events(i).Pages(X)
                        .chkVariable = Buffer.ReadLong
                        .VariableIndex = Buffer.ReadLong
                        .VariableCondition = Buffer.ReadLong
                        .VariableCompare = Buffer.ReadLong
                            
                        .chkSwitch = Buffer.ReadLong
                        .SwitchIndex = Buffer.ReadLong
                        .SwitchCompare = Buffer.ReadLong
                            
                        .chkHasItem = Buffer.ReadLong
                        .HasItemIndex = Buffer.ReadLong
                            
                        .chkSelfSwitch = Buffer.ReadLong
                        .SelfSwitchIndex = Buffer.ReadLong
                        .SelfSwitchCompare = Buffer.ReadLong
                            
                        .GraphicType = Buffer.ReadLong
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .GraphicX2 = Buffer.ReadLong
                        .GraphicY2 = Buffer.ReadLong
                            
                        .MoveType = Buffer.ReadLong
                        .MoveSpeed = Buffer.ReadLong
                        .MoveFreq = Buffer.ReadLong
                            
                        .MoveRouteCount = Buffer.ReadLong
                        
                        .IgnoreMoveRoute = Buffer.ReadLong
                        .RepeatMoveRoute = Buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(mapNum).Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).Index = Buffer.ReadLong
                                .MoveRoute(Y).Data1 = Buffer.ReadLong
                                .MoveRoute(Y).Data2 = Buffer.ReadLong
                                .MoveRoute(Y).Data3 = Buffer.ReadLong
                                .MoveRoute(Y).Data4 = Buffer.ReadLong
                                .MoveRoute(Y).data5 = Buffer.ReadLong
                                .MoveRoute(Y).data6 = Buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = Buffer.ReadLong
                        .DirFix = Buffer.ReadLong
                        .WalkThrough = Buffer.ReadLong
                        .ShowName = Buffer.ReadLong
                        .Trigger = Buffer.ReadLong
                        .CommandListCount = Buffer.ReadLong
                            
                        .Position = Buffer.ReadLong
                    End With
                        
                    If Map(mapNum).Events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map(mapNum).Events(i).Pages(X).CommandList(0 To Map(mapNum).Events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map(mapNum).Events(i).Pages(X).CommandListCount
                            Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount = Buffer.ReadLong
                            Map(mapNum).Events(i).Pages(X).CommandList(Y).ParentList = Buffer.ReadLong
                            If Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map(mapNum).Events(i).Pages(X).CommandList(Y).Commands(1 To Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount)
                                For z = 1 To Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map(mapNum).Events(i).Pages(X).CommandList(Y).Commands(z)
                                        .Index = Buffer.ReadLong
                                        .Text1 = Buffer.ReadString
                                        .Text2 = Buffer.ReadString
                                        .Text3 = Buffer.ReadString
                                        .Text4 = Buffer.ReadString
                                        .Text5 = Buffer.ReadString
                                        .Data1 = Buffer.ReadLong
                                        .Data2 = Buffer.ReadLong
                                        .Data3 = Buffer.ReadLong
                                        .Data4 = Buffer.ReadLong
                                        .data5 = Buffer.ReadLong
                                        .data6 = Buffer.ReadLong
                                        .ConditionalBranch.CommandList = Buffer.ReadLong
                                        .ConditionalBranch.Condition = Buffer.ReadLong
                                        .ConditionalBranch.Data1 = Buffer.ReadLong
                                        .ConditionalBranch.Data2 = Buffer.ReadLong
                                        .ConditionalBranch.Data3 = Buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = Buffer.ReadLong
                                        .MoveRouteCount = Buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).Index = Buffer.ReadLong
                                                .MoveRoute(w).Data1 = Buffer.ReadLong
                                                .MoveRoute(w).Data2 = Buffer.ReadLong
                                                .MoveRoute(w).Data3 = Buffer.ReadLong
                                                .MoveRoute(w).Data4 = Buffer.ReadLong
                                                .MoveRoute(w).data5 = Buffer.ReadLong
                                                .MoveRoute(w).data6 = Buffer.ReadLong
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    'End Event Data

    Call SendMapNpcsToMap(mapNum)
    Call SpawnMapNpcs(mapNum)
    Call SpawnGlobalEvents(mapNum)
    
    For i = 1 To Player_HighIndex
        If Player(i).Map = mapNum Then
            SpawnMapEventsFor i, mapNum
        End If
    Next

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(mapNum)
    Call MapCache_Create(mapNum)
    Call ClearTempTile(mapNum)
    Call CacheResources(mapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            Call PlayerWarp(i, mapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(mapNum)

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SpawnMapEventsFor(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, invNum) < 1 Or GetPlayerInvItemNum(Index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(Index, invNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, invNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (Index)

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    npcNum = Buffer.ReadLong

    ' Prevent hacking
    If npcNum < 0 Or npcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(npcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcNum)
    Call SaveNpc(npcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & npcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = Buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    spellNum = Buffer.ReadLong

    ' Prevent hacking
    If spellNum < 0 Or spellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellNum)
    Call SaveSpell(spellNum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & spellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong 'CLng(Parse(1))
    Y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If X < 0 Or X > Map(GetPlayerMap(Index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        ' Change target
                        If TempPlayer(Index).targetType = TARGET_TYPE_PLAYER And TempPlayer(Index).target = i Then
                            TempPlayer(Index).target = 0
                            TempPlayer(Index).targetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget Index
                        Else
                            TempPlayer(Index).target = i
                            TempPlayer(Index).targetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget Index
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).NPC(i).X = X Then
                If MapNpc(GetPlayerMap(Index)).NPC(i).Y = Y Then
                    If TempPlayer(Index).target = i And TempPlayer(Index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).target = 0
                        TempPlayer(Index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget Index
                    Else
                        ' Change target
                        TempPlayer(Index).target = i
                        TempPlayer(Index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget Index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    PlayerUnequipItem Index, Buffer.ReadByte
    
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Set Buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(Index) < 4 Then Exit Sub
    SetPlayerExp Index, GetPlayerNextLevel(Index)
    CheckPlayerLevelUp Index
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellSlot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellSlot) > GetTickCount Then
        PlayerMsg Index, "Não pode esquecer esse jutsu enquanto não estiver pronto!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellSlot Then
        PlayerMsg Index, "Não pode esquecer esse jutsu enquanto estivar usando!", BrightRed
        Exit Sub
    End If
    
    Player(Index).Spell(spellSlot).Num = 0
    Player(Index).Spell(spellSlot).Level = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopSlot As Long
    Dim shopNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopSlot = Buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(Index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopSlot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        If Player(Index).Money = 0 Or Player(Index).Money < .costvalue Then
            PlayerMsg Index, "Você não tem dinheiro suficiente para isso.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        Player(Index).Money = Player(Index).Money - .costvalue
        Call SendMoney(Index)
        
        GiveInvItem Index, .Item, .ItemValue
    
        ' send confirmation message & reset their shop action
        PlayerMsg Index, "Você acabou de comprar " & Strings.Trim$(Item(.Item).Name), BrightGreen
    End With
    
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invSlot) < 1 Or GetPlayerInvItemNum(Index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(Index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = Item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem Index, itemnum, 1
    GiveInvItem Index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Trade successful.", BrightGreen
    ResetShopAction Index
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem Index, BankSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem Index, invSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim X As Long
    Dim Y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, X
        SetPlayerY Index, Y
        SendPlayerXYToMap Index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).X
    tY = Player(tradeTarget).Y
    sX = Player(Index).X
    sY = Player(Index).Y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    If TempPlayer(Index).InTrade > 0 Then
        TempPlayer(Index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(Index).TradeRequest
        ' let them know they're trading
        PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
        PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(Index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(Index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = Index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade Index, tradeTarget
        SendTrade tradeTarget, Index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate Index, 0
        SendTradeUpdate Index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus Index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                itemnum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem(i).Num = itemnum
                    tmpTradeItem(i).Value = TempPlayer(Index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
                itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).Num = itemnum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).Num > 0 Then
                ' give away!
                GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).Num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory Index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(Index).TradeOffer(i).Num = 0
            TempPlayer(Index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg Index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
        SendCloseTrade Index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(Index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).Value = TempPlayer(Index).TradeOffer(i).Value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).Value > GetPlayerInvItemValue(Index, invSlot) Then
                    TempPlayer(Index).TradeOffer(i).Value = GetPlayerInvItemValue(Index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(Index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Byte, Slot As Byte, hotbarNum As Byte, Actualhotbar As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadByte
    Slot = Buffer.ReadByte
    hotbarNum = Buffer.ReadByte
    Actualhotbar = Buffer.ReadByte
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum, Actualhotbar).Slot = 0
            Player(Index).Hotbar(hotbarNum, Actualhotbar).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(Index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum, Actualhotbar).Slot = Player(Index).Inv(Slot).Num
                        Player(Index).Hotbar(hotbarNum, Actualhotbar).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(Slot).Num > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(Slot).Num).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum, Actualhotbar).Slot = Player(Index).Spell(Slot).Num
                        Player(Index).Hotbar(hotbarNum, Actualhotbar).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Byte
    Dim i As Byte, n As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadByte
    n = Buffer.ReadByte
    
    Select Case Player(Index).Hotbar(Slot, n).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Inv(i).Num = Player(Index).Hotbar(Slot, n).Slot Then
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i).Num > 0 Then
                    If Player(Index).Spell(i).Num = Player(Index).Hotbar(Slot, n).Slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(Index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(Index).target = Index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(Index).target) Or Not IsPlaying(TempPlayer(Index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(Index).target) <> GetPlayerMap(Index) Then Exit Sub
    
    ' init the request
    Party_Invite Index, TempPlayer(Index).target
End Sub

Sub HandleAcceptParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(Index).partyInvite, Index
End Sub

Sub HandleDeclineParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(Index).partyInvite, Index
End Sub

Sub HandlePartyLeave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave Index
End Sub

Sub HandleEventChatReply(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    eventID = Buffer.ReadLong
    pageID = Buffer.ReadLong
    reply = Buffer.ReadLong
    
    If TempPlayer(Index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(Index).EventProcessingCount
            If TempPlayer(Index).EventProcessing(i).eventID = eventID And TempPlayer(Index).EventProcessing(i).pageID = pageID Then
                If TempPlayer(Index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowText Then
                            TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(Index).EventProcessing(i).ListLeftOff(TempPlayer(Index).EventProcessing(i).CurList) = TempPlayer(Index).EventProcessing(i).CurSlot
                                    TempPlayer(Index).EventProcessing(i).CurList = Map(GetPlayerMap(Index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(Index).EventProcessing(i).CurList).Commands(TempPlayer(Index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(Index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(Index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set Buffer = Nothing
End Sub

Sub HandleEvent(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim X As Long, Y As Long, begineventprocessing As Boolean, z As Long, Buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
    End Select
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data
    i = Buffer.ReadLong
    Set Buffer = Nothing
    
    If TempPlayer(Index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(Index).EventMap.CurrentEvents
            If TempPlayer(Index).EventMap.EventPages(z).eventID = i Then
                i = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(Index).EventProcessingCount = TempPlayer(Index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(Index).EventProcessing(TempPlayer(Index).EventProcessingCount)
            With TempPlayer(Index).EventProcessing(TempPlayer(Index).EventProcessingCount)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(Index).EventMap.EventPages(i).eventID
                .pageID = TempPlayer(Index).EventMap.EventPages(i).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(Index)).Events(TempPlayer(Index).EventMap.EventPages(i).eventID).Pages(TempPlayer(Index).EventMap.EventPages(i).pageID).CommandListCount)
            End With
        End If
        begineventprocessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (Index)
End Sub

Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = Buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = Buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set Buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Private Sub HandlePlusSkill(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellSlot As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    spellSlot = Buffer.ReadByte
    Set Buffer = Nothing
    
    If GetPlayerSpellLevel(Index, spellSlot) < MAX_SPELL_LEVEL Then
        If Player(Index).Energy > 0 Then
            Player(Index).Energy = Player(Index).Energy - 1
            Call SetPlayerSpellLevel(Index, spellSlot, GetPlayerSpellLevel(Index, spellSlot) + 1)
            
            SendPlayerSpells Index
        Else
            PlayerMsg Index, "Você não possui energia suficiente para isso!", Orange
        End If
    Else
        PlayerMsg Index, "Esse mágia já está em seu level máximo", Orange
    End If
End Sub

' Solicitar editores
Sub HandleRequestEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ' Obter dados
    Set Buffer = New clsBuffer
    Call Buffer.WriteBytes(Data())

    ' Selecionar editor
    Select Case Buffer.ReadInteger
        Case 0 ' Editor de missões
            Set Buffer = New clsBuffer
            Buffer.WriteLong SQuestEditor
            SendDataTo Index, Buffer.ToArray()
    End Select
    
    ' Limpar
    Set Buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(Quest(n))
    ReDim QuestData(QuestSize - 1)
    QuestData = Buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(QuestData(0)), QuestSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(Index) & " saved Quest #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendQuests Index
End Sub

Sub HandlePlayerHandleQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim QuestNum As Long, Order As Long, i As Long, n As Long
    Dim RemoveStartItems As Boolean
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    QuestNum = Buffer.ReadLong
    Order = Buffer.ReadLong '1 = accept quest, 2 = cancel quest
    
    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If FindOpenInvSlot(Index, Quest(QuestNum).RewardItem(i).Item) = 0 Then
                    PlayerMsg Index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            If FindOpenInvSlot(Index, Quest(QuestNum).GiveItem(i).Item) = 0 Then
                                PlayerMsg Index, "You have no inventory space. Please delete something to take the quest.", BrightRed
                                RemoveStartItems = True
                                Exit For
                            Else
                                GiveInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        If RemoveStartItems = False Then 'this means everything went ok
            Player(Index).PlayerQuest(QuestNum).Status = QUEST_STARTED '1
            Player(Index).PlayerQuest(QuestNum).ActualTask = 1
            Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
            PlayerMsg Index, "New quest accepted: " & Trim$(Quest(QuestNum).Name) & "!", BrightGreen
        End If
        '/alatar v1.2
        
    ElseIf Order = 2 Then
        Player(Index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED '2
        Player(Index).PlayerQuest(QuestNum).ActualTask = 1
        Player(Index).PlayerQuest(QuestNum).CurrentCount = 0
        RemoveStartItems = True 'avoid exploits
        PlayerMsg Index, Trim$(Quest(QuestNum).Name) & " has been canceled!", BrightGreen
    End If
    
    If RemoveStartItems = True Then
        For i = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(i).Item > 0 Then
                If HasItem(Index, Quest(QuestNum).GiveItem(i).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(i).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, Quest(QuestNum).GiveItem(i).Value
                    Else
                        For n = 1 To Quest(QuestNum).GiveItem(i).Value
                            TakeInvItem Index, Quest(QuestNum).GiveItem(i).Item, 1
                        Next
                    End If
                End If
            End If
        Next
    End If
    
    
    SavePlayer Index
    SendPlayerData Index
    SendPlayerQuests Index
    
    Set Buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests Index
End Sub
