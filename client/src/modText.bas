Attribute VB_Name = "modText"
Option Explicit

'The size of a FVF vertex
Private Const FVF_Size As Long = 28

' Local da pasta das fontes
Private Const Path_Font As String = "\data files\graphics\fonts\"

'Point API
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' Estrutura do Caracter
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

' Estrutura do Cabeçalho
Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

' Estrutura da fonte
Private Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
    xOffset As Long
    yOffset As Long
End Type

' Fonts
Public Enum Fonts
    ' Georgia
    Georgia = 1
    
    ' Rockwell
    Rockwell
    
    ' Verdana
    Verdana
    
    ' count value
    Fonts_Count
End Enum

' Store the fonts
Private Font() As CustomFont

' text color pointers
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Gold As Byte = 17
Public Const LightGreen As Byte = 18

' Caracter para alteração de cor
Public Const ColourChar As String * 1 = "#"

' CHATBOX
Public Type ChatStruct
    text As String
    Color As Long
    Visible As Boolean
    Timer As Long
    Channel As Byte
End Type

Private Chat_HighIndex As Byte
Public Const ChatLines As Long = 200
Public Chat(1 To ChatLines) As ChatStruct


' ###############################################
' ## - Propriedades das Fonte - #################

' Carregar fontes
Public Sub LoadFonts()
    'Check if we have the device
    If Direct3D_Device.TestCooperativeLevel <> D3D_OK Then Exit Sub
    ' re-dim the fonts
    ReDim Font(1 To Fonts.Fonts_Count - 1)
    ' load the fonts
    Call SetFont(Fonts.Georgia, "Georgia", 256)
    Call SetFont(Fonts.Rockwell, "Rockwell", 256, 2, 2)
    Call SetFont(Fonts.Verdana, "Verdana", 256)
End Sub

' Setar configurações da fonte
Private Sub SetFont(ByVal fontNum As Long, ByVal texName As String, ByVal size As Long, Optional ByVal xOffset As Long, Optional ByVal yOffset As Long)
    Dim Data() As Byte, f As Long, w As Long, h As Long, Path As String
    
    ' set the path
    Path = App.Path & Path_Font & texName & ".png"
    ' load the texture
    f = FreeFile
    Open Path For Binary As #f
        ReDim Data(0 To LOF(f) - 1)
        Get #f, , Data
    Close #f
    ' get size
    Font(fontNum).TextureSize.X = Data(18) * 256 + Data(19)
    Font(fontNum).TextureSize.Y = Data(22) * 256 + Data(23)
    ' set to struct
    Set Font(fontNum).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, Data(0), UBound(Data) + 1, Font(fontNum).TextureSize.X, Font(fontNum).TextureSize.Y, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    Font(fontNum).xOffset = xOffset
    Font(fontNum).yOffset = yOffset
    ' load header file
    Call LoadFontHeader(Font(fontNum), texName & ".dat")
End Sub

' Carregar cabeçalho da fonte
Private Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal fileName As String)
    Dim FileNum As Byte
    Dim LoopChar As Long
    Dim Row As Single
    Dim u As Single
    Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.Path & Path_Font & fileName For Binary As #FileNum
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

' Obter cores
Public Function DX8Colour(ByVal colourNum As Byte, ByVal alpha As Byte) As Long
    Select Case colourNum
        Case 0 ' Black
            DX8Colour = D3DColorARGB(alpha, 0, 0, 0)
        Case 1 ' Blue
            DX8Colour = D3DColorARGB(alpha, 16, 104, 237)
        Case 2 ' Green
            DX8Colour = D3DColorARGB(alpha, 119, 188, 84)
        Case 3 ' Cyan
            DX8Colour = D3DColorARGB(alpha, 16, 224, 237)
        Case 4 ' Red
            DX8Colour = D3DColorARGB(alpha, 201, 0, 0)
        Case 5 ' Magenta
            DX8Colour = D3DColorARGB(alpha, 255, 0, 255)
        Case 6 ' Brown
            DX8Colour = D3DColorARGB(alpha, 175, 149, 92)
        Case 7 ' Grey
            DX8Colour = D3DColorARGB(alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            DX8Colour = D3DColorARGB(alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            DX8Colour = D3DColorARGB(alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            DX8Colour = D3DColorARGB(alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            DX8Colour = D3DColorARGB(alpha, 157, 242, 242)
        Case 12 ' BrightRed
            DX8Colour = D3DColorARGB(alpha, 255, 0, 0)
        Case 13 ' Pink
            DX8Colour = D3DColorARGB(alpha, 255, 118, 221)
        Case 14 ' Yellow
            DX8Colour = D3DColorARGB(alpha, 255, 255, 0)
        Case 15 ' White
            DX8Colour = D3DColorARGB(alpha, 255, 255, 255)
        Case 16 ' dark brown
            DX8Colour = D3DColorARGB(alpha, 98, 84, 52)
        Case 17 ' gold
            DX8Colour = D3DColorARGB(alpha, 255, 215, 0)
        Case 18 ' light green
            DX8Colour = D3DColorARGB(alpha, 124, 205, 80)
    End Select
End Function

' Obter valor da cor
Function GetColStr(Colour As Long)
    If Colour < 10 Then
        GetColStr = "0" & Colour
    Else
        GetColStr = Colour
    End If
End Function

' Renderizar textos
Public Sub RenderText(ByRef UseFont As Fonts, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Shadow As Boolean = True)
    Dim TempVA(0 To 3) As TLVERTEX, TempStr() As String
    Dim Count As Integer, i As Integer, j As Integer, tmpNum As Integer
    Dim Ascii() As Byte
    Dim TempColor As Long, resetColor As Long
    Dim yOffset As Single
    Dim ignoreChar As Byte

    ' set the color
    Color = DX8Colour(Color, alpha)

    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    resetColor = TempColor
    'Set the texture
    Direct3D_Device.SetTexture 0, Font(UseFont).Texture
    NumTextures = -1
    ' set the position
    X = X - Font(UseFont).xOffset
    Y = Y - Font(UseFont).yOffset
    'Loop through each line if there are line breaks (vbCrLf)
    tmpNum = UBound(TempStr)

    For i = 0 To tmpNum
        If Len(TempStr(i)) > 0 Then
            yOffset = (i * Font(UseFont).CharHeight) + (i * 3)
            Count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            'Loop through the characters
            tmpNum = Len(TempStr(i))
            For j = 1 To tmpNum
                ' check for colour change
                If Mid$(TempStr(i), j, 1) = ColourChar Then
                    Color = Val(Mid$(TempStr(i), j + 1, 2))
                    ' make sure the colour exists
                    If Color = -1 Then
                        TempColor = resetColor
                    Else
                        TempColor = DX8Colour(Color, alpha)
                    End If
                    ignoreChar = 3
                End If
                ' check if we're ignoring this character
                If ignoreChar > 0 Then
                    ignoreChar = ignoreChar - 1
                Else
                    'Copy from the cached vertex array to the temp vertex array
                    Call CopyMemory(TempVA(0), Font(UseFont).HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                    'Set up the verticies
                    TempVA(0).X = X + Count
                    TempVA(0).Y = Y + yOffset
                    TempVA(1).X = TempVA(1).X + X + Count
                    TempVA(1).Y = TempVA(0).Y
                    TempVA(2).X = TempVA(0).X
                    TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                    TempVA(3).X = TempVA(1).X
                    TempVA(3).Y = TempVA(2).Y
                    'Set the colors
                    TempVA(0).Colour = TempColor
                    TempVA(1).Colour = TempColor
                    TempVA(2).Colour = TempColor
                    TempVA(3).Colour = TempColor
                    'Draw the verticies
                    Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                    'Shift over the the position to render the next character
                    Count = Count + Font(UseFont).HeaderInfo.CharWidth(Ascii(j - 1))
                End If
            Next j
        End If
    Next i
End Sub

' Tamanho do texto em pixel
Public Function TextWidth(ByRef UseFont As Fonts, ByVal text As String) As Long
Dim LoopI As Integer, tmpNum As Long, skipCount As Long

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    tmpNum = Len(text)
    For LoopI = 1 To tmpNum
        If Mid$(text, LoopI, 1) = ColourChar Then skipCount = 3
        If skipCount > 0 Then
            skipCount = skipCount - 1
        Else
            TextWidth = TextWidth + Font(UseFont).HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
        End If
    Next LoopI
End Function

' Altura do texto
Public Function TextHeight(ByRef UseFont As Fonts) As Long
    TextHeight = Font(UseFont).HeaderInfo.CellHeight
End Function

' Quebra de texto
Private Function WordWrap(theFont As Fonts, ByVal text As String, ByVal MaxLineLen As Integer, Optional ByRef lineCount As Byte) As String
    Dim TempSplit() As String, TSLoop As Long, lastSpace As Long, size As Long, i As Long, b As Long, tmpNum As Long, skipCount As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    tmpNum = UBound(TempSplit)

    For TSLoop = 0 To tmpNum
        'Clear the values for the new line
        size = 0
        b = 1
        lastSpace = 1

        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine

        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Then
            'Loop through all the characters
            tmpNum = Len(TempSplit(TSLoop))

            For i = 1 To tmpNum
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " "
                        lastSpace = i
                    Case ColourChar
                        skipCount = 3
                End Select
                
                If skipCount > 0 Then
                    skipCount = skipCount - 1
                Else
                    'Add up the size
                    size = size + Font(theFont).HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
                    'Check for too large of a size
                    If size > MaxLineLen Then
                        'Check if the last space was too far back
                        If i - lastSpace > 12 Then
                            'Too far away to the last space, so break at the last character
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                            lineCount = lineCount + 1
                            b = i - 1
                            size = 0
                        Else
                            'Break at the last space to preserve the word
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)) & vbNewLine
                            lineCount = lineCount + 1
                            b = lastSpace + 1
                            'Count all the words we ignored (the ones that weren't printed, but are before "i")
                            size = TextWidth(theFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                        End If
                    End If
    
                    'This handles the remainder
                    If i = Len(TempSplit(TSLoop)) Then
                        If b <> i Then
                            WordWrap = WordWrap & Mid$(TempSplit(TSLoop), b, i)
                            lineCount = lineCount + 1
                        End If
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
    Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, b As Long, tmpNum As Long

    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If

    ' default values
    b = 1
    lastSpace = 1
    size = 0
    tmpNum = Len(text)

    For i = 1 To tmpNum

        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
        End Select

        'Add up the size
        size = size + Font(Fonts.Verdana).HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))

        'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, b, (i - 1) - b))
                b = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, b, lastSpace - b))
                b = lastSpace + 1
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = TextWidth(Fonts.Verdana, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If

        ' Remainder
        If i = Len(text) Then
            If b <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, b, i)
            End If
        End If
    Next
End Sub

' ###################################################################################
' ######### AREA DE DESENHO #########################################################

' CHATBOX
' Adicionar Texto
Public Sub AddText(ByVal text As String, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Channel As Byte = 0)
Dim i As Long

    Chat_HighIndex = 0
    ' Move the rest of it up
    For i = (ChatLines - 1) To 1 Step -1
        If Len(Chat(i).text) > 0 Then
            If i > Chat_HighIndex Then Chat_HighIndex = i + 1
        End If
        Chat(i + 1) = Chat(i)
    Next
    
    Chat(1).text = text
    Chat(1).Color = Color
    Chat(1).Visible = True
    Chat(1).Timer = GetTickCount
    Chat(1).Channel = Channel
End Sub

' Desenhar textos
Public Sub ChatText_Draw()
Dim xO As Long, yO As Long
Dim Colour As Long, yOffset As Long
Dim rLines As Byte, lineCount As Byte
Dim tmpText As String, i As Long, isVisible As Boolean, topWidth As Long, tmpArray() As String, X As Long
    
    ' set the position
    xO = Window(GUI.W_Chatbox).X + 20
    yO = Window(GUI.W_Chatbox).Y + Window(GUI.W_Chatbox).Height - 2
    
    ' loop through chat
    rLines = 1
    i = 1 + ChatScroll
    Do While rLines <= 8
        If i > ChatLines Then Exit Do
        lineCount = 0
        ' exit out early if we come to a blank string
        If Len(Chat(i).text) = 0 Then Exit Do
        ' get visible state
        isVisible = True
        
        'If Options.channelState(Chat(i).Channel) = 0 Then isVisible = False
        ' make sure it's visible
        If isVisible Then
            ' render line
            Colour = Chat(i).Color
            ' check if we need to word wrap
            If TextWidth(Fonts.Verdana, Chat(i).text) > Window(GUI.W_Chatbox).Width - 20 Then
                ' word wrap
                tmpText = WordWrap(Fonts.Verdana, Chat(i).text, Window(GUI.W_Chatbox).Width - 20, lineCount)
                ' can't have it going offscreen.
                If rLines + lineCount > 9 Then Exit Do
                ' continue on
                yOffset = yOffset - (14 * lineCount)
                RenderText Fonts.Verdana, tmpText, xO, yO + yOffset, Colour
                rLines = rLines + lineCount
                ' set the top width
                tmpArray = Split(tmpText, vbNewLine)
                For X = 0 To UBound(tmpArray)
                    If TextWidth(Fonts.Verdana, tmpArray(X)) > topWidth Then topWidth = TextWidth(Fonts.Verdana, tmpArray(X))
                Next
            Else
                ' normal
                yOffset = yOffset - 14
                RenderText Fonts.Verdana, Chat(i).text, xO, yO + yOffset, Colour
                rLines = rLines + 1
                ' set the top width
                If TextWidth(Fonts.Verdana, Chat(i).text) > topWidth Then topWidth = TextWidth(Fonts.Verdana, Chat(i).text)
            End If
        End If
        ' increment chat pointer
        i = i + 1
    Loop
End Sub

' Atualizar textbox do chat
Public Sub UpdateShowChatText()
    Dim i As Long, X As Long
    
    If TextWidth(Fonts.Verdana, MyText) > Window(GUI.W_Chatbox).Width - 65 Then
        For i = Len(MyText) To 1 Step -1
            X = X + Font(Fonts.Verdana).HeaderInfo.CharWidth(Asc(Strings.Mid$(MyText, i, 1)))
            If X > Window(GUI.W_Chatbox).Width - 65 Then
                RenderTextChat = Strings.Right$(MyText, Len(MyText) - i + 1)
                Exit For
            End If
        Next
    Else
        RenderTextChat = MyText
    End If
End Sub

' Desenhar bolha do chat
Public Sub ChatBubble_Draw(ByVal index As Long)
Dim theArray() As String, X As Long, Y As Long, i As Long, MaxWidth As Long, x2 As Long, y2 As Long, Colour As Long
    
    With chatBubble(index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' it's on our map - get co-ords
                X = ConvertMapX((Player(.target).X * 32) + Player(.target).xOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).yOffset) - 50
            End If
        ElseIf .targetType = TARGET_TYPE_NPC Then
            ' it's on our map - get co-ords
            X = ConvertMapX((MapNpc(.target).X * 32) + MapNpc(.target).xOffset) + 16
            Y = ConvertMapY((MapNpc(.target).Y * 32) + MapNpc(.target).yOffset) - 40
        ElseIf .targetType = TARGET_TYPE_EVENT Then
            X = ConvertMapX((Map.MapEvents(.target).X * 32) + Map.MapEvents(.target).xOffset) + 16
            Y = ConvertMapY((Map.MapEvents(.target).Y * 32) + Map.MapEvents(.target).yOffset) - 40
        End If
        
        ' word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
        ' find max width
        For i = 1 To UBound(theArray)
            If TextWidth(Fonts.Verdana, theArray(i)) > MaxWidth Then MaxWidth = TextWidth(Fonts.Verdana, theArray(i))
        Next
                
        ' calculate the new position
        x2 = X - (MaxWidth \ 2)
        y2 = Y - (UBound(theArray) * 12)
                
        ' render bubble - top left
        RenderTexture Tex_ChatBubble, x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, y2 - 5, 111, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_ChatBubble, x2, y2 - 5, 10, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_ChatBubble, x2 - 9, Y, 0, 20, 9, 5, 9, 5
        ' bottom right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, Y, 111, 20, 9, 5, 9, 5
        ' bottom - left half
        RenderTexture Tex_ChatBubble, x2, Y, 10, 20, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_ChatBubble, x2 + (MaxWidth \ 2) + 6, Y, 10, 20, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_ChatBubble, x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, y2, 111, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_ChatBubble, x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_ChatBubble, X - 5, Y, 54, 20, 11, 11, 11, 11
                
        ' render each line centralised
        For i = 1 To UBound(theArray)
            RenderText Fonts.Verdana, theArray(i), X - (TextWidth(Fonts.Verdana, theArray(i)) / 2) - 1, y2 - 2, White
            y2 = y2 + 12
        Next
        ' check if it's timed out - close it if so
        If .Timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

' ##############################################################################
' Nome do jogador
Public Sub PlayerName_Draw(ByVal index As Long)
    Dim textX As Long, textY As Long, Colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the colour
    Colour = White

    If GetPlayerAccess(index) > 0 Then Colour = Yellow ' Admin Color
    
    If GetPlayerPK(index) > 0 Then Colour = BrightRed ' PK Color
    
    textX = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).xOffset + (PIC_X \ 2) - (TextWidth(Fonts.Rockwell, Trim$(GetPlayerName(index))) / 2)
    textY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).yOffset - 16

    If GetPlayerSprite(index) > 0 And GetPlayerSprite(index) <= NumCharacters Then
        textY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).yOffset - (Tex_Character(GetPlayerSprite(index)).Height / 4) + 16
    End If

    ' Desenhar texto
    Call RenderText(Fonts.Rockwell, Trim$(GetPlayerName(index)), textX, textY, Colour)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerName_Draw", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ##############################################################################
' Nome do NPC
Public Sub NpcName_Draw(ByVal index As Long)
    Dim textX As Long, textY As Long, Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Npc(MapNpc(index).Num).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(MapNpc(index).Num).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
        ' get the colour
        If Npc(MapNpc(index).Num).Level <= GetPlayerLevel(MyIndex) - 3 Then
            Colour = Grey
        ElseIf Npc(MapNpc(index).Num).Level <= GetPlayerLevel(MyIndex) - 2 Then
            Colour = Green
        ElseIf Npc(MapNpc(index).Num).Level > GetPlayerLevel(MyIndex) Then
            Colour = Red
        Else
            Colour = White
        End If
    Else
        Colour = White
    End If

    textX = ConvertMapX(MapNpc(index).X * PIC_X) + MapNpc(index).xOffset + (PIC_X \ 2) - (TextWidth(Fonts.Rockwell, Trim$(Npc(MapNpc(index).Num).name)) / 2)
    If Npc(MapNpc(index).Num).Sprite < 1 Or Npc(MapNpc(index).Num).Sprite > NumCharacters Then
        textY = ConvertMapY(MapNpc(index).Y * PIC_Y) + MapNpc(index).yOffset - 16
    Else
        ' Determine location for text
        textY = ConvertMapY(MapNpc(index).Y * PIC_Y) + MapNpc(index).yOffset - (Tex_Character(Npc(MapNpc(index).Num).Sprite).Height / 4) + 16
    End If

    Call RenderText(Fonts.Rockwell, Trim$(Npc(MapNpc(index).Num).name), textX, textY, Colour)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcName_Draw", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ##############################################################################
' Atributos do editor de mapas
Public Function MapAttributes_Draw()
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
                        
                        ' Desenhar quadrado
                        If .Type > 0 Then RenderTexture Tex_Selection, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 0, 0, 32, 32, 32, 32, DX8Colour(BrightBlue, 200)
                        
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Fonts.Verdana, "B", tx, ty, BrightRed
                            Case TILE_TYPE_WARP
                                RenderText Fonts.Verdana, "W", tx, ty, BrightBlue
                            Case TILE_TYPE_ITEM
                                RenderText Fonts.Verdana, "I", tx, ty, White
                            Case TILE_TYPE_NPCAVOID
                                RenderText Fonts.Verdana, "N", tx, ty, White
                            Case TILE_TYPE_KEY
                                RenderText Fonts.Verdana, "K", tx, ty, White
                            Case TILE_TYPE_KEYOPEN
                                RenderText Fonts.Verdana, "O", tx, ty, White
                            Case TILE_TYPE_RESOURCE
                                RenderText Fonts.Verdana, "B", tx, ty, Green
                            Case TILE_TYPE_DOOR
                                RenderText Fonts.Verdana, "D", tx, ty, Brown
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Fonts.Verdana, "S", tx, ty, Yellow
                            Case TILE_TYPE_SHOP
                                RenderText Fonts.Verdana, "S", tx, ty, BrightBlue
                            Case TILE_TYPE_BANK
                                RenderText Fonts.Verdana, "B", tx, ty, Blue
                            Case TILE_TYPE_HEAL
                                RenderText Fonts.Verdana, "H", tx, ty, BrightGreen
                            Case TILE_TYPE_TRAP
                                RenderText Fonts.Verdana, "T", tx, ty, BrightRed
                            Case TILE_TYPE_SLIDE
                                RenderText Fonts.Verdana, "S", tx, ty, BrightCyan
                            Case TILE_TYPE_SOUND
                                RenderText Fonts.Verdana, "S", tx, ty, White
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "MapAttributes_Draw", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
