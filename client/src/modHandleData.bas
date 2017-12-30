Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    
    'Events
    HandleDataSub(SSpawnEvent) = GetAddress(AddressOf HandleSpawnEventPage)
    HandleDataSub(SEventMove) = GetAddress(AddressOf HandleEventMove)
    HandleDataSub(SEventDir) = GetAddress(AddressOf HandleEventDir)
    HandleDataSub(SEventChat) = GetAddress(AddressOf HandleEventChat)
    
    HandleDataSub(SEventStart) = GetAddress(AddressOf HandleEventStart)
    HandleDataSub(SEventEnd) = GetAddress(AddressOf HandleEventEnd)
    
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    
    HandleDataSub(SMapEventData) = GetAddress(AddressOf HandleMapEventData)
    
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    frmLoad.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picMain.Visible = True
    
    Msg = buffer.ReadString 'Parse(1)
    
    Set buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, Options.Game_Name)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' save options
    Options.SavePass = frmMenu.chkPass.Value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    Player_HighIndex = buffer.ReadLong
    
    Set buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Receiving game data...")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = buffer.ReadLong
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picRegister.Visible = False
    frmLoad.Visible = False
    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).name)
    Next

    frmMenu.cmbClass.ListIndex = 0
    n = frmMenu.cmbClass.ListIndex + 1
    
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            Z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To Z)
            ' loop-receive data
            For X = 0 To Z
                .FemaleSprite(X) = buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InGame = True
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    frmMain.picCharacter.Refresh
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        'frmMain.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
        frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        ' hp bar
        frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        'frmMain.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
        frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        ' mp bar
        frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat index, i, buffer.ReadLong
        frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, buffer.ReadLong
    PlayerXP = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerLevel(i, buffer.ReadLong)
    Call SetPlayerPOINTS(i, buffer.ReadLong)
    Call SetPlayerSprite(i, buffer.ReadLong)
    Call SetPlayerMap(i, buffer.ReadLong)
    Call SetPlayerX(i, buffer.ReadLong)
    Call SetPlayerY(i, buffer.ReadLong)
    Call SetPlayerDir(i, buffer.ReadLong)
    Call SetPlayerAccess(i, buffer.ReadLong)
    Call SetPlayerPK(i, buffer.ReadLong)
    Call SetPlayerClass(i, buffer.ReadLong)
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadLong
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        
        ' Set the character windows
        frmMain.lblCharName = GetPlayerName(MyIndex) & " - Level " & GetPlayerLevel(MyIndex)
        
        For X = 1 To Stats.Stat_Count - 1
            frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X)
        Next
        
        ' Set training label visiblity depending on points
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                If GetPlayerStat(index, X) < 255 Then
                    frmMain.lblTrainStat(X).Visible = True
                Else
                    frmMain.lblTrainStat(X).Visible = False
                End If
            Next
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X).Visible = False
            Next
        End If
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).yOffset = PIC_Y
        Case DIR_DOWN
            Player(i).yOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).xOffset = PIC_X
        Case DIR_RIGHT
            Player(i).xOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(i, Dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong

    With MapNpc(i)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).X = 0
        Blood(i).Y = 0
        Blood(i).Sprite = 0
        Blood(i).Timer = 0
    Next
    
    Map.CurrentEvents = 0
    ReDim Map.MapEvents(0)
    
    ' Get map num
    X = buffer.ReadLong
    ' Get revision
    Y = buffer.ReadLong

    If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
            CacheNewMapSounds
            initAutotiles
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    GettingMap = True
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long, Z As Long, w As Long
Dim buffer As clsBuffer
Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.BGS = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    
    Map.Fog = buffer.ReadLong
    Map.FogSpeed = buffer.ReadLong
    Map.FogOpacity = buffer.ReadLong
    
    Map.Red = buffer.ReadLong
    Map.Green = buffer.ReadLong
    Map.Blue = buffer.ReadLong
    Map.alpha = buffer.ReadLong
    
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            For Z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Autotile(Z) = buffer.ReadLong
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadString
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
        Map.NpcSpawnType(X) = buffer.ReadLong
        n = n + 1
    Next

    ClearTempTile
    initAutotiles
    
    Set buffer = Nothing
    
    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    CacheNewMapSounds

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    
    ' re-position the map name
    Call UpdateDrawMapName
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next
    
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    initAutotiles
    
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.alpha

    GettingMap = False
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(Msg, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapNpc(n)
        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    
    Npc_HighIndex = 0
    
    ' Get the npc high Index
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i
            Exit For
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapKey(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    n = buffer.ReadByte
    TempTile(X, Y).DoorOpen = n
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim SpellNum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SpellNum = buffer.ReadLong
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Update the spells on the pic
    Set buffer = New clsBuffer
    buffer.WriteLong CSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i).Num = buffer.ReadInteger
        PlayerSpells(i).Level = buffer.ReadByte
        Player(MyIndex).Energy = buffer.ReadByte
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = buffer.ReadByte
            MapResource(i).X = buffer.ReadLong
            MapResource(i).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    Call DrawPing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Message As String, Color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    Message = buffer.ReadString
    Color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    CreateActionMsg Message, Color, tmpType, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .Timer = GetTickCount
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNpcNum = buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadLong
    PlayerSpells(Slot).Cooldown = GetTickCount
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Byte
Dim name As String
Dim Message As String
Dim Colour As Long
Dim Header As String
Dim PK As Byte, Channel As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    name = buffer.ReadString
    Access = buffer.ReadByte
    PK = buffer.ReadByte
    Message = buffer.ReadString
    Header = buffer.ReadString
    Colour = buffer.ReadLong
    
    ' Check access level
    Colour = White

    If Access > 0 Then Colour = Yellow ' Cor do administrador
    
    If PK > 0 Then Colour = BrightRed ' Cor do assassino
    
    ' find channel
    Channel = 0
    Select Case Header
        Case "[M]"
            Channel = 1
        Case "[G] "
            Channel = 2
    End Select
    
    ' remove the colour char from the message
    ' Message = Replace$(Message, ColourChar, vbNullString)
    ' add to the chat box
    AddText ColourChar & GetColStr(Colour) & Header & name & ": " & ColourChar & GetColStr(Grey) & Message, Grey, , Channel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopnum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    StunDuration = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = buffer.ReadLong
        Bank.Item(i).Value = buffer.ReadLong
    Next
    
    InBank = True
    frmMain.picCover.Visible = True
    frmMain.picBank.Visible = True
    frmMain.picBank.Refresh
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InTrade = buffer.ReadLong
    frmMain.picCover.Visible = True
    frmMain.picTrade.Visible = True
    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InTrade = 0
    frmMain.picCover.Visible = False
    frmMain.picTrade.Visible = False
    frmMain.lblTradeStatus.Caption = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    dataType = buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = buffer.ReadLong
            TradeYourOffer(i).Value = buffer.ReadLong
        Next
        frmMain.lblYourWorth.Caption = buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = buffer.ReadLong
            TradeTheirOffer(i).Value = buffer.ReadLong
        Next
        frmMain.lblTheirWorth.Caption = buffer.ReadLong & "g"
    End If

    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    tradeStatus = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case tradeStatus
        Case 0 ' clear
            frmMain.lblTradeStatus.Caption = vbNullString
        Case 1 ' they've accepted
            frmMain.lblTradeStatus.Caption = "Other player has accepted."
        Case 2 ' you've accepted
            frmMain.lblTradeStatus.Caption = "Waiting for other player to accept."
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Byte, b As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
        
    For b = 1 To MAX_HOTBAR_NUM
        For i = 1 To MAX_HOTBAR
            Hotbar(i, b).Slot = buffer.ReadByte
            Hotbar(i, b).sType = buffer.ReadByte
        Next
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Player_HighIndex = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    theName = buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    theName = buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    inParty = buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        Next
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
        If Party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = Trim$(GetPlayerName(Party.Member(i)))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    Party.MemberCount = buffer.ReadLong
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim playerNum As Long, partyIndex As Long
Dim buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = buffer.ReadLong
        Player(playerNum).Vital(i) = buffer.ReadLong
    Next
    
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' hp bar
    frmMain.imgPartyHealth(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    ' spr bar
    frmMain.imgPartySpirit(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnEventPage(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long, i As Long, Z As Long, X As Long, Y As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    id = buffer.ReadLong
    
    If id > Map.CurrentEvents Then
        Map.CurrentEvents = id
        ReDim Preserve Map.MapEvents(Map.CurrentEvents)
    End If

    With Map.MapEvents(id)
        .name = buffer.ReadString
        .Dir = buffer.ReadLong
        .ShowDir = .Dir
        .GraphicNum = buffer.ReadLong
        .GraphicType = buffer.ReadLong
        .GraphicX = buffer.ReadLong
        .GraphicX2 = buffer.ReadLong
        .GraphicY = buffer.ReadLong
        .GraphicY2 = buffer.ReadLong
        .MovementSpeed = buffer.ReadLong
        .Moving = 0
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Position = buffer.ReadLong
        .Visible = buffer.ReadLong
        .WalkAnim = buffer.ReadLong
        .DirFix = buffer.ReadLong
        .WalkThrough = buffer.ReadLong
        .ShowName = buffer.ReadLong
    End With
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnEventPage", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim id As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long, ShowDir As Long
Dim Movement As Long, MovementSpeed As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    id = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    ShowDir = buffer.ReadLong
    MovementSpeed = buffer.ReadLong
    
    If id > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(id)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 1
        .ShowDir = ShowDir
        .MovementSpeed = MovementSpeed
        

        Select Case Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    
    If i > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(i)
        .Dir = Dir
        .ShowDir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventChat(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim buffer As clsBuffer
Dim choices As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    frmMain.picEventChat.Visible = True
    EventReplyID = buffer.ReadLong
    EventReplyPage = buffer.ReadLong
    frmMain.lblEventChat.Caption = buffer.ReadString
    frmMain.picEventChat.Visible = True
    frmMain.lblEventChat.Visible = True
    choices = buffer.ReadLong
    
    InEvent = True
    
    For i = 1 To 4
        frmMain.lblChoices(i).Visible = False
    Next
    
    frmMain.lblEventChatContinue.Visible = False
    
    If choices = 0 Then
        frmMain.lblEventChatContinue.Visible = True
    Else
        For i = 1 To choices
            frmMain.lblChoices(i).Visible = True
            frmMain.lblChoices(i).Caption = buffer.ReadString
        Next
    End If
    
    AnotherChat = buffer.ReadLong
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventChat", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventStart(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InEvent = True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventStart", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEventEnd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InEvent = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEventEnd", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayBGM(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString
    
    StopMusic
    PlayMusic str
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString

    PlaySound str, -1, -1
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never
    StopMusic
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStopSound(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 0 To UBound(Sounds()) - 1
        StopSound (i)
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapEventData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, i As Long, X As Long, Y As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    'Event Data!
    Map.EventCount = buffer.ReadLong
        
    If Map.EventCount > 0 Then
        ReDim Map.Events(0 To Map.EventCount)
        For i = 1 To Map.EventCount
            With Map.Events(i)
                .name = buffer.ReadString
                .Global = buffer.ReadLong
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .pageCount = buffer.ReadLong
            End With
            If Map.Events(i).pageCount > 0 Then
                ReDim Map.Events(i).Pages(0 To Map.Events(i).pageCount)
                For X = 1 To Map.Events(i).pageCount
                    With Map.Events(i).Pages(X)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                            
                        .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                            
                        .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                            
                        .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                            
                        .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                            
                        .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                            
                        .MoveRouteCount = buffer.ReadLong
                        
                        .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map.Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).index = buffer.ReadLong
                                .MoveRoute(Y).Data1 = buffer.ReadLong
                                .MoveRoute(Y).Data2 = buffer.ReadLong
                                .MoveRoute(Y).Data3 = buffer.ReadLong
                                .MoveRoute(Y).Data4 = buffer.ReadLong
                                .MoveRoute(Y).Data5 = buffer.ReadLong
                                .MoveRoute(Y).Data6 = buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .WalkThrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                            
                        .Position = buffer.ReadLong
                    End With
                        
                    If Map.Events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map.Events(i).Pages(X).CommandList(0 To Map.Events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map.Events(i).Pages(X).CommandListCount
                            Map.Events(i).Pages(X).CommandList(Y).CommandCount = buffer.ReadLong
                            Map.Events(i).Pages(X).CommandList(Y).ParentList = buffer.ReadLong
                            If Map.Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map.Events(i).Pages(X).CommandList(Y).Commands(1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount)
                                For Z = 1 To Map.Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map.Events(i).Pages(X).CommandList(Y).Commands(Z)
                                        .index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .Data1 = buffer.ReadLong
                                        .Data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .Data5 = buffer.ReadLong
                                        .Data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.Data1 = buffer.ReadLong
                                        .ConditionalBranch.Data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).index = buffer.ReadLong
                                                .MoveRoute(w).Data1 = buffer.ReadLong
                                                .MoveRoute(w).Data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).Data5 = buffer.ReadLong
                                                .MoveRoute(w).Data6 = buffer.ReadLong
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
    
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapEventData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleChatBubble(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, targetType As Long, target As Long, Message As String, Colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    target = buffer.ReadLong
    targetType = buffer.ReadLong
    Message = buffer.ReadString
    Colour = buffer.ReadLong
    
    AddChatBubble target, targetType, Message, Colour
    Set buffer = Nothing
errorhandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpecialEffect(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, effectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    effectType = buffer.ReadLong
    
    Select Case effectType
        Case EFFECT_TYPE_FADEIN
            FadeType = 1
            FadeAmount = 0
        Case EFFECT_TYPE_FADEOUT
            FadeType = 0
            FadeAmount = 255
        Case EFFECT_TYPE_FLASH
            FlashTimer = GetTickCount + 150
        Case EFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    Set buffer = Nothing
errorhandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

