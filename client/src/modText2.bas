Attribute VB_Name = "modText2"
Option Explicit
' Stuffs
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

Sub EngineInitFontTextures()
    ' FONT DEFAULT
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.Path & FONT_PATH & "texdefault.png"
    LoadTexture Font_Default.Texture
    
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.Path & FONT_PATH & "georgia.png"
    LoadTexture Font_Georgia.Texture
End Sub

Sub UnloadFontTextures()
    UnloadFont Font_Georgia
End Sub
Sub UnloadFont(Font As CustomFont)
    Font.Texture.Texture = 0
End Sub

Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal fileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.Path & FONT_PATH & fileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Colour = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            .Vertex(1).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            .Vertex(2).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            .Vertex(3).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Public Sub DrawPlayerName(ByVal index As Long)
Dim textX As Long
Dim textY As Long
Dim Color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(index) = NO Then

        Select Case GetPlayerAccess(index)
            Case 0
                'Color = Orange
            Case 1
                Color = White
            Case 2
                Color = Cyan
            Case 3
                Color = BrightGreen
            Case 4
                Color = Yellow
        End Select

    Else
        Color = BrightRed
    End If

    name = Trim$(Player(index).name)
    ' calc pos
    textX = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).xOffset + (PIC_X \ 2) - (TextWidth(Fonts.Verdana, (Trim$(name))) / 2)
    If GetPlayerSprite(index) < 1 Or GetPlayerSprite(index) > NumCharacters Then
        textY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).yOffset - 16
    Else
        ' Determine location for text
        textY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).yOffset - (Tex_Character(GetPlayerSprite(index)).Height / 4) + 16
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Fonts.Verdana, name, textX, textY, Color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal index As Long)
Dim textX As Long
Dim textY As Long
Dim Color As Long
Dim name As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    npcNum = MapNpc(index).Num

    Select Case Npc(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            Color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            Color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            Color = Grey
        Case Else
            Color = BrightGreen
    End Select

    name = Trim$(Npc(npcNum).name)
    textX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).xOffset + (PIC_X \ 2) - (TextWidth(Fonts.Verdana, (Trim$(name))) / 2)
    If Npc(npcNum).Sprite < 1 Or Npc(npcNum).Sprite > NumCharacters Then
        textY = ConvertMapY(MapNpc(index).Y * PIC_Y) + MapNpc(index).yOffset - 16
    Else
        ' Determine location for text
        textY = ConvertMapY(MapNpc(index).Y * PIC_Y) + MapNpc(index).yOffset - (Tex_Character(Npc(npcNum).Sprite).Height / 4) + 16
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Fonts.Verdana, name, textX, textY, Color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function DrawMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tx As Long
    Dim ty As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        tx = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Fonts.Verdana, "B", tx, ty, BrightRed, 0
                            Case TILE_TYPE_WARP
                                RenderText Fonts.Verdana, "W", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_ITEM
                                RenderText Fonts.Verdana, "I", tx, ty, White, 0
                            Case TILE_TYPE_NPCAVOID
                                RenderText Fonts.Verdana, "N", tx, ty, White, 0
                            Case TILE_TYPE_KEY
                                RenderText Fonts.Verdana, "K", tx, ty, White, 0
                            Case TILE_TYPE_KEYOPEN
                                RenderText Fonts.Verdana, "O", tx, ty, White, 0
                            Case TILE_TYPE_RESOURCE
                                RenderText Fonts.Verdana, "B", tx, ty, Green, 0
                            Case TILE_TYPE_DOOR
                                RenderText Fonts.Verdana, "D", tx, ty, Brown, 0
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Fonts.Verdana, "S", tx, ty, Yellow, 0
                            Case TILE_TYPE_SHOP
                                RenderText Fonts.Verdana, "S", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_BANK
                                RenderText Fonts.Verdana, "B", tx, ty, Blue, 0
                            Case TILE_TYPE_HEAL
                                RenderText Fonts.Verdana, "H", tx, ty, BrightGreen, 0
                            Case TILE_TYPE_TRAP
                                RenderText Fonts.Verdana, "T", tx, ty, BrightRed, 0
                            Case TILE_TYPE_SLIDE
                                RenderText Fonts.Verdana, "S", tx, ty, BrightCyan, 0
                            Case TILE_TYPE_SOUND
                                RenderText Fonts.Verdana, "S", tx, ty, White, 0
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawActionMsg(ByVal index As Long)
    Dim X As Long, Y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(index).Y > 0 Then
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) - 2
            Else
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(index).Y > 0 Then
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            Else
                X = ActionMsg(index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> index Then
                        ClearActionMsg index
                        index = i
                    End If
                End If
            Next
            X = (frmMain.Width \ 2) - ((Len(Trim$(ActionMsg(index).Message)) \ 2) * 8)
            Y = 425

    End Select
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(index).Created + Time Then
        RenderText Fonts.Verdana, ActionMsg(index).Message, X, Y, ActionMsg(index).Color, 0
    Else
        ClearActionMsg index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEventName(ByVal index As Long)
Dim textX As Long
Dim textY As Long
Dim Color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If InMapEditor Then Exit Sub

    Color = White

    name = Trim$(Map.MapEvents(index).name)
    
    ' calc pos
    textX = ConvertMapX(Map.MapEvents(index).X * PIC_X) + Map.MapEvents(index).xOffset + (PIC_X \ 2) - (TextWidth(Fonts.Verdana, (Trim$(name))) / 2)
    If Map.MapEvents(index).GraphicType = 0 Then
        textY = ConvertMapY(Map.MapEvents(index).Y * PIC_Y) + Map.MapEvents(index).yOffset - 16
    ElseIf Map.MapEvents(index).GraphicType = 1 Then
        If Map.MapEvents(index).GraphicNum < 1 Or Map.MapEvents(index).GraphicNum > NumCharacters Then
            textY = ConvertMapY(Map.MapEvents(index).Y * PIC_Y) + Map.MapEvents(index).yOffset - 16
        Else
            ' Determine location for text
            textY = ConvertMapY(Map.MapEvents(index).Y * PIC_Y) + Map.MapEvents(index).yOffset - (Tex_Character(Map.MapEvents(index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(index).GraphicType = 2 Then
        If Map.MapEvents(index).GraphicY2 > 0 Then
            textY = ConvertMapY(Map.MapEvents(index).Y * PIC_Y) + Map.MapEvents(index).yOffset - ((Map.MapEvents(index).GraphicY2 - Map.MapEvents(index).GraphicY) * 32) + 16
        Else
            textY = ConvertMapY(Map.MapEvents(index).Y * PIC_Y) + Map.MapEvents(index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Fonts.Verdana, name, textX, textY, Color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
