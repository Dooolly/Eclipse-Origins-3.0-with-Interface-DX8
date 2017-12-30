Attribute VB_Name = "modWindow"
Option Explicit

' Local do mouse
Public MouseX As Single
Public MouseY As Single

' Janela que será movida
Private MovableWind As Byte

' ChatBox
Public Const MAX_LINES As Byte = 200
Public InChat As Boolean
Public ChatCursor As String
Public RenderTextChat As String
Public MyText As String
Public ChatScroll As Long
Public totalChatLines As Long
Public ChatButtonUp As Boolean
Public ChatButtonDown As Boolean

' Descrição
Private ActualItemDesc As Long
Private TYPE_DESC As Byte

' Constantes das janelas
Public Enum GUI
   W_None
   W_MainBar
   W_Inventory
   W_HUD
   W_DescriptionItem
   W_Chatbox
   W_Spells

   ' Quantidade de Janelas
   W_Count
End Enum

' Estrutura do Botão
Private Type ButtonRec
    X As Single
    Y As Single
    Pos_X As Single
    Pos_Y As Single
    Width As Long
    Height As Long
    State As Byte
    Visible As Boolean
    text As String
End Type

' Estrutura da Janela
Private Type WindowRec
    Visible As Boolean
    Movable As Boolean
    X As Long
    Y As Long
    Width As Long
    Height As Long
    Buttons() As ButtonRec
End Type

' Variavel onde ficará a imagem principal
Private Tex_GUI As DX8TextureRec

' Variavel das Janelas
Public Window(1 To GUI.W_Count - 1) As WindowRec

' Carregar configurações padrão das janelas
Public Sub InitGUI()
    Dim i As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Carregar Textura
    ReDim Preserve gTexture(NumTextures + 1)
    Tex_GUI.filepath = App.Path & "\data files\graphics\gui\gui_texture.png"
    Tex_GUI.Texture = NumTextures + 1
    LoadTexture Tex_GUI
    
    ActualHotbar = 1
    ChatScroll = 0
    
    ' Carregar Barra principal
    With Window(GUI.W_MainBar)
        .X = 202
        .Y = 547
        .Width = 620
        .Height = 53
        .Visible = True
        .Movable = False
        ReDim .Buttons(1 To 9)
        
        For i = 1 To 7
            .Buttons(i).X = .X + 364 + (36 * (i - 1))
            .Buttons(i).Y = .Y + 15
            .Buttons(i).Width = 35
            .Buttons(i).Height = 35
            .Buttons(i).Pos_X = 1 + (71 * (i - 1))
            .Buttons(i).Pos_Y = 496
        Next
        
        .Buttons(8).X = .X + 347
        .Buttons(8).Y = .Y + 16
        .Buttons(8).Width = 12
        .Buttons(8).Height = 12
        .Buttons(8).Pos_X = 498
        .Buttons(8).Pos_Y = 507
        
        .Buttons(9).X = .X + 347
        .Buttons(9).Y = .Y + 39
        .Buttons(9).Width = 12
        .Buttons(9).Height = 12
        .Buttons(9).Pos_X = 511
        .Buttons(9).Pos_Y = 507
    End With
    
    ' Carregar Inventário
    With Window(GUI.W_Inventory)
        .X = 250
        .Y = 200
        .Width = 245
        .Height = 216
        .Visible = False
        .Movable = True
    End With
    
    ' Interface do jogador
    With Window(GUI.W_HUD)
        .X = 1
        .Y = 1
        .Width = 199
        .Height = 66
        .Visible = True
        .Movable = False
    End With
    
    ' Carregar ItemDesc
    With Window(GUI.W_DescriptionItem)
        .X = 0
        .Y = 0
        .Width = 156
        .Height = 64
        .Visible = False
        .Movable = False
    End With
    
    ' Carregar Mágias
    With Window(GUI.W_Spells)
        .X = 300
        .Y = 189
        .Width = 143
        .Height = 216
        .Visible = False
        .Movable = True
        ReDim Preserve .Buttons(1 To MAX_PLAYER_SPELLS)
        For i = 1 To MAX_PLAYER_SPELLS
            With .Buttons(i)
                .X = 0
                .Y = 0
                .Pos_X = 391
                .Pos_Y = 1
                .Width = 14
                .Height = 14
                .State = 0
            End With
        Next
    End With
    
    ' Carregar CHAT
    With Window(GUI.W_Chatbox)
        .X = 1
        .Y = 400
        .Width = 300
        .Height = 113
        .Visible = True
        .Movable = False
        ReDim .Buttons(1 To 6)
        
        For i = 1 To UBound(.Buttons)
            .Buttons(i).X = .X
            .Buttons(i).Y = .Y + (19 * (i - 1))
            .Buttons(i).Width = 18
            .Buttons(i).Height = 18
            .Buttons(i).Pos_Y = 363 + (19 * (i - 1))
            .Buttons(i).Pos_X = 1
        Next
        
        .Buttons(1).State = 2
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitGUI", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Desenhar toda a interface
Public Sub DrawInterface()
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call Draw_MainBar ' Desenhar barra principal
    Call HUD_Draw ' Desenhar interface
    Call Chatbox_Draw
    
    If Window(GUI.W_Inventory).Visible Then Call Inventory_Draw ' Desenhar inventário
    If Window(GUI.W_Spells).Visible Then Call Skills_Draw ' Desenhar janela de magias
    If Window(GUI.W_DescriptionItem).Visible Then Call DrawItemDesc ' Descrição do item
    
    If DragInvSlotNum > 0 Then Call DrawDragItem
    If DragSpell > 0 Then Call DrawDragSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInterface", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função do botão do mouse está em movimento
Public Sub MouseMove_Handle(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MouseX = X
    MouseY = Y

    ' Checar qual janela o mouse está
    If Not InMapEditor Then
        For i = 1 To GUI.W_Count - 1
            If (X >= Window(i).X And X <= Window(i).X + Window(i).Width) And (Y >= Window(i).Y And Y <= Window(i).Y + Window(i).Height) Then
                If Window(i).Visible Then
                    Select Case i
                        Case GUI.W_MainBar
                            ' none
                        Case GUI.W_Inventory
                            Call Inventory_MouseMove(Button, X, Y)
                    End Select
                End If
            End If
        Next
    End If

    ' Checar qual janela o mouse está
    If Button = 1 Then
        If MovableWind > 0 Then
            Window(MovableWind).X = Window(MovableWind).X + (MouseX - SOffsetX)
            Window(MovableWind).Y = Window(MovableWind).Y + (MouseY - SOffsetY)
            SOffsetX = X
            SOffsetY = Y
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MouseMove_Handle", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função do botão do mouse está em baixo
Public Sub MouseDown_Handle(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SOffsetX = X
    SOffsetY = Y

    ' Checar qual janela o mouse está
    If Not InMapEditor Then
        For i = 1 To GUI.W_Count - 1
            If (X >= Window(i).X And X <= Window(i).X + Window(i).Width) And (Y >= Window(i).Y And Y <= Window(i).Y + Window(i).Height) Then
                If Window(i).Visible Then
                    ' Fechar janela selecionada
                    If (X >= (Window(i).X + Window(i).Width - 22) And X <= Window(i).X + Window(i).Width) And (Y >= Window(i).Y And Y <= Window(i).Y + 22) Then
                        Window(i).Visible = False
                        Exit Sub
                    End If
                    
                    ' Definir qual a janela que será movida
                    If (X >= Window(i).X And X <= Window(i).X + Window(i).Width - 22) And (Y >= Window(i).Y And Y <= Window(i).Y + 22) Then
                        If Window(i).Visible And Window(i).Movable Then MovableWind = i
                    End If
                    
                    ' Executar função de cada janela
                    Select Case i
                        Case GUI.W_MainBar
                            MainBar_MouseDown Button, X, Y
                            Exit Sub
                        Case GUI.W_Inventory
                            Call Inventory_MouseDown(Button, X, Y)
                            Exit Sub
                        Case GUI.W_Chatbox
                            Chat_MouseDown Button, X, Y
                            Exit Sub
                        Case GUI.W_Spells
                            Call Skills_MouseDown(Button, X, Y)
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' left click
    If Button = vbLeftButton Then
        ' targetting
        'Call PlayerSearch(CurX, CurY, False)
        
    ' right click
    ElseIf Button = vbRightButton Then
        If ShiftDown Then
            ' admin warp if we're pressing shift and right clicking
            If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
        Else
            ' targetting
            'Call PlayerSearch(CurX, CurY, True)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MouseDown_Handle", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função do botão do mouse está em cima
Public Sub MouseUp_Handle(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, n As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checar qual janela o mouse está
    If Not InMapEditor Then
        For i = 1 To GUI.W_Count - 1
            If (X >= Window(i).X And X <= Window(i).X + Window(i).Width) And (Y >= Window(i).Y And Y <= Window(i).Y + Window(i).Height) Then
                If Window(i).Visible Then
                    Select Case i
                        Case GUI.W_MainBar
                            MainBar_MouseUp Button, X, Y
                            GoTo ClearSub
                        Case GUI.W_Inventory
                            Call Inventory_MouseUp(Button, X, Y)
                            GoTo ClearSub
                    End Select
                End If
            End If
        Next
    End If
           
    ' Dropar itens no chão
    If DragInvSlotNum > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, DragInvSlotNum)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerInvItemValue(MyIndex, DragInvSlotNum) > 0 Then
                CurrencyMenu = 1 ' drop
                frmMain.lblCurrency.Caption = "Quantos items deseja jogar fora?"
                tmpCurrencyItem = DragInvSlotNum
                frmMain.txtCurrency.text = vbNullString
                frmMain.picCurrency.Visible = True
                frmMain.txtCurrency.SetFocus
            End If
        Else
            Call SendDropItem(DragInvSlotNum, 0)
        End If
    End If
    
ClearSub:
    
    ' Limpar variaveis
    DragInvSlotNum = 0
    DragSpell = 0
    MovableWind = 0
    For i = 1 To GUI.W_Count - 1
        Select Case i
            Case GUI.W_None, GUI.W_Inventory, GUI.W_HUD, GUI.W_DescriptionItem
                ' Aqui as janelas que não contenham botões
            Case Else
                For n = 1 To UBound(Window(i).Buttons)
                    If i = GUI.W_Chatbox Then
                        Window(i).Buttons(5).State = 0
                        Window(i).Buttons(6).State = 0
                        Exit For
                    Else
                        Window(i).Buttons(n).State = 0
                    End If
                Next
        End Select
    Next
    
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MouseUp_Handle", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função do botão do duplo clique
Public Sub DblClick_Handle()
    Dim i As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checar qual janela o mouse está
    If Not InMapEditor Then
        For i = 1 To GUI.W_Count - 1
            If (MouseX >= Window(i).X And MouseX <= Window(i).X + Window(i).Width) And (MouseY >= Window(i).Y And MouseY <= Window(i).Y + Window(i).Height) Then
                If Window(i).Visible Then
                    Select Case i
                        Case GUI.W_Inventory
                            Call Inventory_DblClick(MouseX, MouseY)
                        Case GUI.W_Spells
                            Call Skills_DblClick
                    End Select
                End If
            End If
        Next
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DblClick_Handle", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ##############################################################
' ### - Barra Principal - ######################################
Private Sub Draw_MainBar()
    Dim dRect As RECT, sRect As RECT
    Dim Num As String, i As Byte, n As Byte
    Dim Amount As Long, Colour As Long
    Dim XP_Width As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Window(GUI.W_MainBar)
        ' Desenha designer
        RenderTexture Tex_GUI, .X, .Y, 1, 547, .Width, .Height, .Width, .Height
        
        ' Desenhar barra de EXP
        XP_Width = ((GetPlayerExp(MyIndex) / .Width) / (PlayerXP / .Width)) * .Width
        RenderTexture Tex_GUI, .X, .Y, 1, 532, XP_Width, 14, XP_Width, 14
        RenderText Fonts.Verdana, Int(GetPlayerExp(MyIndex) / PlayerXP * 100) & "%", .X + (.Width / 2) - (TextWidth(Fonts.Verdana, Int(GetPlayerExp(MyIndex) / PlayerXP * 100)) / 2), .Y, White, 200
        
        ' Loop para exibir itens
        For i = 1 To MAX_HOTBAR
            With dRect
                .Top = Window(GUI.W_MainBar).Y + 17
                .Left = Window(GUI.W_MainBar).X + ((1 + 33) * (((i - 1) Mod 10))) + 6
                .Bottom = .Top + 32
                .Right = .Left + 32
            End With
        
            With sRect
                .Top = 0
                .Left = 0
                .Bottom = 32
                .Right = 32
            End With
        
            Select Case Hotbar(i, ActualHotbar).sType
                Case 1 ' Inventário
                    If Len(Item(Hotbar(i, ActualHotbar).Slot).name) > 0 Then
                        If Item(Hotbar(i, ActualHotbar).Slot).Pic > 0 Then
                            If Item(Hotbar(i, ActualHotbar).Slot).Pic <= numitems Then
                                RenderTextureByRects Tex_Item(Item(Hotbar(i, ActualHotbar).Slot).Pic), sRect, dRect
                            End If
                        End If
                    End If
                Case 2 ' Magia
                    If Len(Spell(Hotbar(i, ActualHotbar).Slot).name) > 0 Then
                        If Spell(Hotbar(i, ActualHotbar).Slot).Icon > 0 Then
                            If Spell(Hotbar(i, ActualHotbar).Slot).Icon <= NumSpellIcons Then
                                ' Checar se a magia está congelada
                                For n = 1 To MAX_PLAYER_SPELLS
                                    If PlayerSpells(n).Num = Hotbar(i, ActualHotbar).Slot Then
                                        ' has spell
                                        If Not PlayerSpells(i).Cooldown = 0 Then
                                            sRect.Left = 32
                                            sRect.Right = 64
                                        End If
                                    End If
                                Next
                                
                                RenderTextureByRects Tex_SpellIcon(Spell(Hotbar(i, ActualHotbar).Slot).Icon), sRect, dRect
                            End If
                        End If
                    End If
            End Select
        
            ' Renderizar Letras
            Num = i
            If Strings.Trim$(Num) = "10" Then Num = "0"
            RenderText Fonts.Verdana, Strings.Trim$(Num), dRect.Left + 2, dRect.Top + 19, White
        Next
        
        ' Desenha Hotbar Atual
        RenderText Fonts.Verdana, ActualHotbar, .X + 349, .Y + 26, White, 200
        
        ' Loop para desenhar botões
        For i = 1 To 9
            With .Buttons(i)
                If i < 8 Then
                    If .State = 2 Then ' Clicado
                        RenderTexture Tex_GUI, .X, .Y + 1, .Pos_X + .Width, .Pos_Y, .Width, .Height, .Width, .Height
                    ElseIf (MouseX >= .X And MouseX <= .X + .Width) And (MouseY >= .Y And MouseY <= .Y + .Height) Then
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y, .Width, .Height, .Width, .Height
                        ' play sound
                        If Not LastButtonSound_Main = i Then
                            PlaySound Sound_ButtonHover, -1, -1
                            LastButtonSound_Main = i
                        End If
                    End If
                Else
                    If .State = 2 Then ' Clicado
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y + .Width, .Width, .Height, .Width, .Height
                    ElseIf (MouseX >= .X And MouseX <= .X + .Width) And (MouseY >= .Y And MouseY <= .Y + .Height) Then
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y, .Width, .Height, .Width, .Height
                        ' play sound
                        If Not LastButtonSound_Main = i Then
                            PlaySound Sound_ButtonHover, -1, -1
                            LastButtonSound_Main = i
                        End If
                    End If
                End If
            End With
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Draw_MainBar", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função do Hotbar apagar ou usar item e magia
Private Sub MainBar_MouseDown(Button As Integer, X As Single, Y As Single)
    Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find out which button we're clicking
    For SlotNum = 1 To UBound(Window(GUI.W_MainBar).Buttons)
        With Window(GUI.W_MainBar).Buttons(SlotNum)
            ' check if we're on the button
            If (X >= .X And X <= .X + .Width) And (Y >= .Y And Y <= .Y + .Height) Then
                .State = 2 ' clicked
                Select Case SlotNum
                    Case 1 ' Inventário
                        Window(GUI.W_Inventory).Visible = Not Window(GUI.W_Inventory).Visible
                    Case 2 ' Status
                         
                    Case 3 ' Magias
                        Window(GUI.W_Spells).Visible = Not Window(GUI.W_Spells).Visible
                    Case 8 ' Mudar de Hotbar - Aumentar
                         If ActualHotbar >= 3 Then
                            ActualHotbar = 1
                         Else
                            ActualHotbar = ActualHotbar + 1
                         End If
                    Case 9 ' Mudar de Hotbar
                         If ActualHotbar <= 1 Then
                            ActualHotbar = 3
                         Else
                            ActualHotbar = ActualHotbar - 1
                         End If
                End Select
                
                ' Executar Som
                PlaySound Sound_ButtonClick, -1, -1
                Exit Sub
            End If
        End With
    Next
    
    ' Hotbar
    SlotNum = IsHotbarSlot(X, Y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MainBar_MouseDown", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        Top = Window(GUI.W_MainBar).Y + 17
        Left = Window(GUI.W_MainBar).X + ((1 + 33) * (((i - 1) Mod 10))) + 6
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub MainBar_MouseUp(Button As Integer, X As Single, Y As Single)
    Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)
    If SlotNum = 0 Then Exit Sub
    
    ' Items
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, SlotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' Magia
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, SlotNum
        DragSpell = 0
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MainBar_MouseUp", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ###########################################################################################
' ## Desenhar inventário #####################################
Private Sub Inventory_Draw()
    Dim InvSlot As Byte, itemNum As Integer, itemPic As Byte
    Dim Rec As RECT, Rec_Pos As RECT
    Dim Amount As Long, Colour As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Window(GUI.W_Inventory)
        ' Desenhar janela
        RenderTexture Tex_GUI, .X, .Y, 1, 1, .Width, .Height, .Width, .Height
        
        ' Renderizar ouro do jogador
        'If GetPlayerGold(MyIndex) > 0 Then
            'RenderText Font_Normal, Strings.Format(GetPlayerGold(MyIndex), "###,###,###"), .X + 47, .Y + 205, White
        'Else
            RenderText Fonts.Verdana, "Dinheiro: 0", .X + 6, .Y + 198, White
        'End If
        
        ' Loop Para desenhar os items
        For InvSlot = 1 To MAX_INV
            itemNum = GetPlayerInvItemNum(MyIndex, InvSlot)

            ' Se o item for maior que 0 ou menor ou igual Máximo de items
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                itemPic = Item(itemNum).Pic

                ' Checar se o item tem um icone
                If itemPic > 0 And itemPic <= numitems Then
                    With Rec
                        .Top = 0
                        .Left = 0
                        .Bottom = 32
                        .Right = 32
                    End With

                    With Rec_Pos
                        .Top = Window(GUI.W_Inventory).Y + 25 + ((2 + 32) * ((InvSlot - 1) \ 7))
                        .Bottom = .Top + 32
                        .Left = Window(GUI.W_Inventory).X + 4 + ((2 + 32) * (((InvSlot - 1) Mod 7)))
                        .Right = .Left + 32
                    End With

                    RenderTextureByRects Tex_Item(itemPic), Rec, Rec_Pos

                    ' Se for acumulativo
                    If GetPlayerInvItemValue(MyIndex, InvSlot) > 1 Then
                        Amount = GetPlayerInvItemValue(MyIndex, InvSlot)
                        
                        ' Desenhar de acordo com K M B.
                        If Amount < 1000000 Then
                            Colour = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            Colour = Yellow
                        ElseIf Amount > 10000000 Then
                            Colour = BrightGreen
                        End If
                        
                        ' Desenhar quantidade
                        RenderText Fonts.Verdana, Strings.Format(ConvertCurrency(Amount), "#,###,###,###"), Rec_Pos.Left, Rec_Pos.Top + 19, Colour, 0
                    End If
                End If
            End If
        Next
    End With
    
' Error handler
    Exit Sub
errorhandler:
    HandleError "Inventory_Draw", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' É um item do inventário
Private Function IsInvItem(ByVal X As Single, ByVal Y As Single, Optional ByVal EmptySlot As Boolean = False) As Long
    Dim TempRec As RECT, SkipThis As Boolean
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV
        If Not EmptySlot Then
            If GetPlayerInvItemNum(MyIndex, i) <= 0 Or GetPlayerInvItemNum(MyIndex, i) > MAX_ITEMS Then SkipThis = True
        End If

        If Not SkipThis Then
            With TempRec
                .Top = Window(GUI.W_Inventory).Y + 25 + ((2 + 32) * ((i - 1) \ 7))
                .Bottom = .Top + PIC_Y
                .Left = Window(GUI.W_Inventory).X + 4 + ((2 + 32) * (((i - 1) Mod 7)))
                .Right = .Left + PIC_X
            End With
    
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If
        SkipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' Função mousedown do inventário
Private Sub Inventory_MouseDown(Button As Integer, X As Single, Y As Single)
    Dim invNum As Long
    
    ' Caso o debug tiver ativo, mostrar se houve algum erro
    If Options.Debug = 1 Then On Error GoTo errorhandler

    invNum = IsInvItem(X, Y)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Then Exit Sub
            DragInvSlotNum = invNum
        End If
    End If
    
    Window(GUI.W_DescriptionItem).Visible = False
    ActualItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Inventory_MouseDown", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Inventory_MouseUp(Button As Integer, X As Single, Y As Single)
    Dim InvSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InTrade > 0 Then Exit Sub
    If InBank Then Exit Sub

    If DragInvSlotNum > 0 Then
        InvSlot = IsInvItem(X, Y, True)
        If InvSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, InvSlot
    End If

    DragInvSlotNum = 0
    Window(GUI.W_DescriptionItem).Visible = False
    ActualItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Inventory_MouseUp", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Inventory_MouseMove(Button As Integer, X As Single, Y As Single)
    Dim InvSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If DragInvSlotNum > 0 Then Exit Sub
    
    InvSlot = IsInvItem(X, Y, True)
    
    If InvSlot <> 0 Then
        If GetPlayerInvItemNum(MyIndex, InvSlot) > 0 Then
            ActualItemDesc = InvSlot
            TYPE_DESC = 0
            With Window(GUI.W_DescriptionItem)
                .X = Window(GUI.W_Inventory).X - .Width - 1
                .Y = Window(GUI.W_Inventory).Y
                .Visible = True
            End With
        Else
            Window(GUI.W_DescriptionItem).Visible = False
            ActualItemDesc = 0
        End If
    Else
        Window(GUI.W_DescriptionItem).Visible = False
        ActualItemDesc = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Inventory_MouseMove", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Função de duplo clique
Private Sub Inventory_DblClick(X As Single, Y As Single)
    Dim invNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    invNum = IsInvItem(X, Y)

    If invNum <> 0 Then
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CONSUME Then
                CurrencyMenu = 2 ' deposit
                frmMain.lblCurrency.Caption = "Quantos itens você quer depositar?"
                tmpCurrencyItem = invNum
                frmMain.txtCurrency.text = vbNullString
                frmMain.picCurrency.Visible = True
                frmMain.txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).Num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)).Type = ITEM_TYPE_CONSUME Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CONSUME Then
                CurrencyMenu = 4 ' offer in trade
                frmMain.lblCurrency.Caption = "Quantos você quer trocar?"
                tmpCurrencyItem = invNum
                frmMain.txtCurrency.text = vbNullString
                frmMain.picCurrency.Visible = True
                frmMain.txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Inventory_DblClick", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemNum > 0 Then Exit Sub
    
    PicNum = Item(itemNum).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub
    RenderTexture Tex_Item(PicNum), MouseX - 16, MouseY - 16, 0, 0, 32, 32, 32, 32
End Sub

' ##############################################################
' ### - HUD Player - ###########################################
Private Sub HUD_Draw()
    Dim Data As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Window(GUI.W_HUD)
        ' Desenhar Background das Barras
        RenderTexture Tex_GUI, .X, .Y, 415, 1, .Width, .Height, .Width, .Height
        ' Desenhar face
        If GetPlayerSprite(MyIndex) <= UBound(Tex_Face) Then RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), .X + 2, .Y + 2, 0, 0, 62, 62, 62, 62
        ' Desenhar Nome do Jogador
        RenderText Fonts.Verdana, Strings.Trim$(GetPlayerName(MyIndex)), .X + 70, .Y + 3, White
        ' Desenhar Level do jogador
        RenderText Fonts.Verdana, "Lv. " & GetPlayerLevel(MyIndex), .X + 68, .Y + 47, White
        
        ' Desenhar Barra de HP
        Data = ((GetPlayerVital(MyIndex, HP) / 124) / (GetPlayerMaxVital(MyIndex, HP) / 124)) * 124
        If Data > 0 Then RenderTexture Tex_GUI, .X + 66, .Y + 17, 415, 68, Data, 16, Data, 16
        ' Desenhar HP atual/Max
        RenderText Fonts.Verdana, GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP), .X + 66 + (124 / 2) - (TextWidth(Fonts.Verdana, GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)) / 2), .Y + 18, White, 70
        
        ' Desenhar Barra de MP
        Data = ((GetPlayerVital(MyIndex, MP) / 118) / (GetPlayerMaxVital(MyIndex, MP) / 118)) * 118
        If Data > 0 Then RenderTexture Tex_GUI, .X + 66, .Y + 32, 415, 83, Data, 13, Data, 13
        ' Desenhar MP atual/Max
        RenderText Fonts.Verdana, GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP), .X + 66 + (118 / 2) - (TextWidth(Fonts.Verdana, GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)) / 2), .Y + 32, White, 100
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HUD_Draw", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ##############################################################
' ### - DESCRIÇÃO DO ITEM - ####################################
Private Sub DrawItemDesc()
    Dim PicNum As Long
    Dim Colour As Long
    Dim Height As Byte
    Dim tmpText As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With Window(GUI.W_DescriptionItem)
        ' Desenha designer
        RenderTexture Tex_GUI, .X, .Y, 1, 218, .Width, 20, .Width, 20
        RenderTexture Tex_GUI, .X, .Y + 20, 1, 223, .Width, 13, .Width, 11
        
        ' Desenhar dados do item
        If ActualItemDesc > 0 Then
            If TYPE_DESC = 1 Then
                PicNum = GetPlayerEquipment(MyIndex, ActualItemDesc)
            Else
                PicNum = GetPlayerInvItemNum(MyIndex, ActualItemDesc)
            End If

            If PicNum > 0 Then
                ' Raridade do item
                Select Case Item(PicNum).Rarity
                    Case 1
                        Colour = BrightGreen
                    Case 2
                        Colour = BrightBlue
                    Case 3
                        Colour = Magenta
                    Case 4
                        Colour = Yellow
                    Case 5
                        'Colour = Orange
                    Case Else
                        Colour = White
                End Select
                
                ' Nome do item
                tmpText = Strings.Trim$(Item(PicNum).name)
                RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 5, Colour

                ' Aprimoramento do item
                'If GetPlayerInvItemCombine(MyIndex, ActualItemDesc) > 0 Then
                    'RenderText Fonts.Verdana, "+" & GetPlayerInvItemCombine(MyIndex, ActualItemDesc), .X + .Width - 23, .Y + 5, Colour
                'End If

                ' Classe Requerida
                If Item(PicNum).ClassReq > 0 Then
                    RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                    Height = Height + 1
                    
                    tmpText = Strings.Trim$(Class(Item(PicNum).ClassReq).name)
                    If GetPlayerClass(MyIndex) = Item(PicNum).ClassReq Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
                    
                    RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Colour
                End If

                ' Level Requerido
                If Item(PicNum).LevelReq > 0 Then
                    RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                    Height = Height + 1
                    
                    tmpText = "Level: " & Item(PicNum).LevelReq
                    If GetPlayerLevel(MyIndex) >= Item(PicNum).LevelReq Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
             
                    RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Colour
                End If
                
                ' Valor
                If Item(PicNum).Price > 0 Then
                    RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                    Height = Height + 1
                    
                    tmpText = "Valor: " & Item(PicNum).Price & "g"
                    Colour = Yellow
             
                    RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Colour
                End If
                
                ' Tipo do Item
                Select Case Item(PicNum).Type
                    Case ITEM_TYPE_WEAPON
                        tmpText = "Arma"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100

                        ' Add Stats
                        For Colour = 1 To Stats.Stat_Count - 1
                            If Item(PicNum).Add_Stat(Colour) > 0 Then
                                RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                                Height = Height + 1
                    
                                Select Case Colour
                                    Case Stats.Strength
                                        tmpText = "Força: +"
                                    Case Stats.Endurance
                                        tmpText = "Defesa: +"
                                    Case Stats.Intelligence
                                        tmpText = "Inteligencia: +"
                                    Case Stats.Agility
                                        tmpText = "Agilidade: +"
                                    Case Stats.Willpower
                                        tmpText = "Recuperação: +"
                                End Select
                    
                                tmpText = tmpText & Item(PicNum).Add_Stat(Colour)
                                RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), White
                            End If
                        Next
                    Case ITEM_TYPE_ARMOR
                        tmpText = "Armadura"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                        
                        ' Add Stats
                        For Colour = 1 To Stats.Stat_Count - 1
                            If Item(PicNum).Add_Stat(Colour) > 0 Then
                                RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                                Height = Height + 1
                    
                                Select Case Colour
                                    Case Stats.Strength
                                        tmpText = "Força: +"
                                    Case Stats.Endurance
                                        tmpText = "Defesa: +"
                                    Case Stats.Intelligence
                                        tmpText = "Inteligencia: +"
                                    Case Stats.Agility
                                        tmpText = "Agilidade: +"
                                    Case Stats.Willpower
                                        tmpText = "Recuperação: +"
                                End Select
                    
                                tmpText = tmpText & Item(PicNum).Add_Stat(Colour)
                                RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), White
                            End If
                        Next
                    Case ITEM_TYPE_HELMET
                        tmpText = "Capacete"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                        
                        ' Add Stats
                        For Colour = 1 To Stats.Stat_Count - 1
                            If Item(PicNum).Add_Stat(Colour) > 0 Then
                                RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                                Height = Height + 1
                    
                                Select Case Colour
                                    Case Stats.Strength
                                        tmpText = "Força: +"
                                    Case Stats.Endurance
                                        tmpText = "Defesa: +"
                                    Case Stats.Intelligence
                                        tmpText = "Inteligencia: +"
                                    Case Stats.Agility
                                        tmpText = "Agilidade: +"
                                    Case Stats.Willpower
                                        tmpText = "Recuperação: +"
                                End Select
                    
                                tmpText = tmpText & Item(PicNum).Add_Stat(Colour)
                                RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), White
                            End If
                        Next
                    Case ITEM_TYPE_SHIELD
                        tmpText = "Acessório"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                        
                        ' Add Stats
                        For Colour = 1 To Stats.Stat_Count - 1
                            If Item(PicNum).Add_Stat(Colour) > 0 Then
                                RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                                Height = Height + 1
                    
                                Select Case Colour
                                    Case Stats.Strength
                                        tmpText = "Força: +"
                                    Case Stats.Endurance
                                        tmpText = "Defesa: +"
                                    Case Stats.Intelligence
                                        tmpText = "Inteligencia: +"
                                    Case Stats.Agility
                                        tmpText = "Agilidade: +"
                                    Case Stats.Willpower
                                        tmpText = "Recuperação: +"
                                End Select
                    
                                tmpText = tmpText & Item(PicNum).Add_Stat(Colour)
                                RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), White
                            End If
                        Next
                    Case ITEM_TYPE_CONSUME
                        tmpText = "Consumo"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                    
                        ' Exibir o total de HP que recupera
                        If Item(PicNum).AddHP > 0 Then
                            RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                            Height = Height + 1
                            
                            tmpText = "Recuperar HP: " & Item(PicNum).AddHP
                            RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Pink
                        End If
                        
                        ' Exibir o total de MP que recupera
                        If Item(PicNum).AddMP > 0 Then
                            RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                            Height = Height + 1
                            
                            tmpText = "Recuperar MP: " & Item(PicNum).AddMP
                            RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Pink
                        End If
                        
                        ' Exibir o total de EXP que ganha
                        If Item(PicNum).AddEXP > 0 Then
                            RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                            Height = Height + 1
                            
                            tmpText = "Adicionar EXP: " & Item(PicNum).AddEXP
                            RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 32 + (13 * (Height - 1)), Pink
                        End If
                    Case ITEM_TYPE_CURRENCY
                        tmpText = "Item de Missão"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                    
                        ' Desenha descrição
                        If Strings.Trim$(Item(PicNum).Desc) <> vbNullString Then
                            Dim Text_Array() As String
                            Dim Text_Lines(0 To 9) As String
                            Dim i As Byte
                            
                            Text_Array = Split(Strings.Trim(Item(PicNum).Desc), " ")

                            For Colour = 0 To UBound(Text_Array)
                                If TextWidth(Fonts.Verdana, Strings.Trim$(Text_Lines(i) & Text_Array(Colour) & " ")) < .Width - 12 Then
                                    Text_Lines(i) = Text_Lines(i) & Text_Array(Colour) & " "
                                Else
                                    i = i + 1
                                    If i > 9 Then
                                        i = 9
                                        Exit For
                                    End If
                                    Text_Lines(i) = Text_Lines(i) & Text_Array(Colour) & " "
                                End If
                            Next
                            
                            For Colour = 0 To i
                                RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 223, .Width, 13, .Width, 11
                                Height = Height + 1
                                RenderText Fonts.Verdana, Strings.Trim$(Text_Lines(Colour)), .X + (.Width / 2) - (TextWidth(Fonts.Verdana, Strings.Trim$(Text_Lines(Colour))) / 2), .Y + 32 + (13 * (Height - 1)), White
                            Next
                        End If
                    Case ITEM_TYPE_SPELL
                        tmpText = "Mágia"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                    'Case ITEM_TYPE_PRECIOUS
                        'tmpText = "Precioso"
                        'RenderText Font_Normal, tmpText, .X + (.Width / 2) - (TextWidth(Font_Normal, tmpText) / 2), .Y + 19, White, 100
                    'Case ITEM_TYPE_STONE
                        'tmpText = "Pedra Mágica"
                        'RenderText Font_Normal, tmpText, .X + (.Width / 2) - (TextWidth(Font_Normal, tmpText) / 2), .Y + 19, White, 100
                    Case Else
                        tmpText = "Nenhum"
                        RenderText Fonts.Verdana, tmpText, .X + (.Width / 2) - (TextWidth(Fonts.Verdana, tmpText) / 2), .Y + 19, White, 100
                End Select
            End If
        End If
        
        ' Desenha designer
        RenderTexture Tex_GUI, .X, .Y + 33 + (13 * Height), 1, 238, .Width, 5, .Width, 5
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ####################################################################################
' ### MAGIAS #################################################
Public Sub DrawDragSpell()
    Dim PicNum As Integer, SpellNum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    SpellNum = PlayerSpells(DragSpell).Num
    If Not SpellNum > 0 Then Exit Sub
    
    PicNum = Spell(SpellNum).Icon

    If PicNum < 1 Or PicNum > NumSpellIcons Then Exit Sub

    RenderTexture Tex_SpellIcon(PicNum), MouseX - 16, MouseY - 16, 0, 0, 32, 32, 32, 32
End Sub

Private Sub Skills_Draw() ' Desenha Hotbar
    Dim spellslot As Byte
    Dim SpellNum As Long, SpellIcon As Long
    Dim Rec As RECT, Rec_Pos As RECT
    Dim BarWidth_GUIEnergy As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Window(GUI.W_Spells)
        ' Desenha designer
        RenderTexture Tex_GUI, .X, .Y, 247, 1, .Width, .Height, .Width, .Height
        
        ' Desenha energia
        RenderText Fonts.Verdana, "Energia: " & Player(MyIndex).Energy, .X + 6, .Y + 198, White

        'If Player(MyIndex).Energy_EXP > 0 Then
            ' health bar
            'BarWidth_GUIEnergy = ((Player(MyIndex).Energy_EXP / 137) / (MAX_SPELL_EXP / 137)) * 137
            
            ' Desenha energia bar
            'RenderTexture Tex_GUI, .X + 5, .Y + 188, 254, 202, BarWidth_GUIEnergy, 10, BarWidth_GUIEnergy, 10
        'End If
        
        ' Desenhar Skills
        For spellslot = 1 To MAX_PLAYER_SPELLS
            SpellNum = PlayerSpells(spellslot).Num

            If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
                SpellIcon = Spell(SpellNum).Icon

                If SpellIcon > 0 And SpellIcon <= NumSpellIcons Then
                    With Rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
                    
                    ' Checar se está com cooldown
                    If Not PlayerSpells(spellslot).Cooldown = 0 Then
                        Rec.Left = 32
                        Rec.Right = 64
                    End If
                    
                    ' Checar se tem level para ser usada
                    If PlayerSpells(spellslot).Level <= 0 Then
                        Rec.Left = 32
                        Rec.Right = 64
                    End If
                    
                    ' Verificar nivel de mana
                    If Spell(SpellNum).MPCost > Player(MyIndex).Vital(Vitals.MP) Then
                        Rec.Left = 32
                        Rec.Right = 64
                    End If
                    
                    ' Checar se a magia precisa se arma e o jogador não tem
                    'If Spell(SpellNum).WeapReq Then
                        'If GetPlayerEquipment(MyIndex, Weapon) <= 0 Then
                            'Rec.Left = 33
                            'Rec.Right = 66
                        'End If
                    'End If

                    With Rec_Pos
                        .Top = Window(GUI.W_Spells).Y + 26 + ((1 + 32) * ((spellslot - 1) \ 4))
                        .Bottom = .Top + PIC_Y
                        .Left = Window(GUI.W_Spells).X + 5 + ((1 + 32) * (((spellslot - 1) Mod 4)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_SpellIcon(SpellIcon), Rec, Rec_Pos
                    ' Desenha Level
                    RenderText Fonts.Verdana, PlayerSpells(spellslot).Level, Rec_Pos.Left, Rec_Pos.Top + 19, White
                End If
                
                ' Desenha botão de Upgrade
                If Player(MyIndex).Energy > 0 And PlayerSpells(spellslot).Level < MAX_SPELL_LEVEL Then
                    With .Buttons(spellslot)
                        .X = Rec_Pos.Left + 18
                        .Y = Rec_Pos.Top + 18
                        
                        If .State = 2 Then ' Clicado
                            RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y + (.Width * 2), .Width, .Height, .Width, .Height
                        ElseIf (MouseX >= .X And MouseX <= .X + .Width) And (MouseY >= .Y And MouseY <= .Y + .Height) Then
                            RenderTexture Tex_GUI, .X, .Y - 1, .Pos_X, .Pos_Y + .Width, .Width, .Height, .Width, .Height
                            ' play sound
                            If Not LastButtonSound_Main = spellslot Then
                                PlaySound Sound_ButtonHover, -1, -1
                                LastButtonSound_Main = spellslot
                            End If
                        Else ' Normal
                            RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y, .Width, .Height, .Width, .Height
                        End If
                    End With
                End If
            End If
        Next
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Skills_Draw", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single, Optional ByVal EmptySlot As Boolean = False) As Byte
Dim TempRec As RECT, SkipThis As Boolean
Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS
        If Not EmptySlot Then
            If PlayerSpells(i).Num <= 0 And PlayerSpells(i).Num > MAX_SPELLS Then SkipThis = True
        End If

        If Not SkipThis Then
            With TempRec
                .Top = Window(GUI.W_Spells).Y + 26 + ((4 + 32) * ((i - 1) \ 4))
                .Bottom = .Top + PIC_Y
                .Left = Window(GUI.W_Spells).X + 5 + ((4 + 32) * (((i - 1) Mod 4)))
                .Right = .Left + PIC_X
            End With
    
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If
        
        SkipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub Skills_MouseDown(Button As Integer, X As Single, Y As Single)
    Dim SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellNum = IsPlayerSpell(X, Y)
    If Button = 1 Then ' left click
        If SpellNum <> 0 Then
            'If Player(MyIndex).Energy > 0 And PlayerSpells(SpellNum).Level < MAX_SPELL_LEVEL Then
                With Window(GUI.W_Spells).Buttons(SpellNum)
                    If (MouseX >= .X And MouseX <= .X + .Width) And (MouseY >= .Y And MouseY <= .Y + .Height) Then
                        .State = 2
                        Send_PlusSpell SpellNum
                        Exit Sub
                    End If
                End With
            'End If
            
            If PlayerSpells(SpellNum).Level > 0 Then
                DragSpell = SpellNum
                Exit Sub
            End If
        End If
    ElseIf Button = 2 Then ' right click
        If SpellNum <> 0 Then
            If PlayerSpells(SpellNum).Num > 0 Then
                Dialogue "Deletar Magia", "Tem certeza de que quer deletar a magia " & Strings.Trim$(Spell(PlayerSpells(SpellNum).Num).name) & "?", DIALOGUE_TYPE_FORGET, True, SpellNum
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Skills_MouseDown", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Skills_DblClick()
     Dim SpellNum As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellNum = IsPlayerSpell(MouseX, MouseX)

    If SpellNum <> 0 Then
        Call CastSpell(SpellNum)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Skills_DblClick", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ##############################################################
' ### - CHAT - #################################################
Private Sub Chatbox_Draw()
    Dim i As Long, tmpY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Window(GUI.W_Chatbox)
        ' Desenha designer
        RenderTexture Tex_GUI, .X + 19, .Y, 39, 363, .Width - 19, 2, 16, 2
        RenderTexture Tex_GUI, .X + 19, .Y, 39, 365, .Width - 19, .Height, 16, 73
        
        ' Desenha barra para escrever
        If InChat Then
            RenderTexture Tex_GUI, .X, .Y + (.Height + 1), 1, 477, .Width, 18, .Width, 18
            RenderText Fonts.Verdana, RenderTextChat & ChatCursor, .X + 5, .Y + .Height + 3, White
        End If
        
        ' Desenhar textos so chat
        Call ChatText_Draw
        
        ' Desenhar botões
        For i = 1 To UBound(.Buttons)
            With .Buttons(i)
                If i <= 4 Then
                    If .State = 2 Then ' Clicado
                        RenderTexture Tex_GUI, .X + 1, .Y, .Pos_X + .Width, .Pos_Y, .Width, .Height, .Width, .Height
                    Else ' Normal
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y, .Width, .Height, .Width, .Height
                    End If
                Else
                    If .State = 2 Then ' Clicado
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X + (.Width * 2), .Pos_Y, .Width, .Height, .Width, .Height
                    ElseIf (MouseX >= .X And MouseX <= .X + .Width) And (MouseY >= .Y And MouseY <= .Y + .Height) Then
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X + .Width, .Pos_Y, .Width, .Height, .Width, .Height
                    Else ' Normal
                        RenderTexture Tex_GUI, .X, .Y, .Pos_X, .Pos_Y, .Width, .Height, .Width, .Height
                    End If
                End If
            End With
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawCHAT", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' scroll bar
Private Sub Chat_MouseDown(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' find out which button we're clicking
    For i = 1 To UBound(Window(GUI.W_Chatbox).Buttons)
        With Window(GUI.W_Chatbox).Buttons(i)
            ' check if we're on the button
            If (X >= .X And X <= .X + .Width) And (Y >= .Y And Y <= .Y + .Height) Then
                Select Case i
                    Case 1, 2, 3, 4
                        MsgPrivate = vbNullString
                        MyText = vbNullString
                        UpdateShowChatText
                        Window(GUI.W_Chatbox).Buttons(1).State = 0
                        Window(GUI.W_Chatbox).Buttons(2).State = 0
                        Window(GUI.W_Chatbox).Buttons(3).State = 0
                        Window(GUI.W_Chatbox).Buttons(4).State = 0
                        .State = 2
                    Case 5 ' up
                        .State = 2 ' clicked
                        ChatButtonUp = True
                    Case 6 ' down
                        .State = 2 ' clicked
                        ChatButtonDown = True
                End Select
            End If
        End With
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Chat_MouseDown", "modWindow", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
