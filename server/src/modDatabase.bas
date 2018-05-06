Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

    If ServerLog Then
        filename = App.path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.path & "\data\options.ini", "OPTIONS", "Website")
    
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.path & "data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & "Server"
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Integer, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal Sprite As Integer)
    Dim F As Long
    Dim n As Byte

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Class = ClassNum
        
        If Player(Index).Sex = SEX_MALE Then
            Player(Index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(Index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(Index).Level = 1
        Player(Index).Energy = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(Index).Dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).X = START_X
        Player(Index).Y = START_Y
        Player(Index).Dir = DIR_DOWN
        Player(Index).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(Index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(Index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).Name)) > 0 Then
                        Player(Index).Spell(n).Num = Class(ClassNum).StartSpell(n)
                        Player(Index).Spell(n).Num = Class(ClassNum).StartSpellLevel(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Append As #F
            Print #F, Name
        Close #F
        
        Call SavePlayer(Index)
    End If
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If
    Next
End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim F As Long

    filename = App.path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
    Call ClearPlayer(Index)
    filename = App.path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Class = 1

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If
End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim X As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        Class(i).startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        
        If Class(i).startItemCount > 0 And Class(i).startItemCount <= MAX_INV Then
            ReDim Class(i).StartItem(1 To Class(i).startItemCount)
            ReDim Class(i).StartValue(1 To Class(i).startItemCount)
        
            ' loop for items & values
            For X = 1 To Class(i).startItemCount
                Class(i).StartItem(X) = Val(GetVar(filename, "CLASS" & i, "StartItem" & X))
                Class(i).StartValue(X) = Val(GetVar(filename, "CLASS" & i, "StartValue" & X))
            Next
        End If
        
        ' loop for spells
        Class(i).startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        
        ' how many starting spells?
        If Class(i).startSpellCount > 0 And Class(i).startSpellCount <= MAX_INV Then
            ReDim Class(i).StartSpell(1 To Class(i).startSpellCount)
            ReDim Class(i).StartSpellLevel(1 To Class(i).startSpellCount)
            
            For X = 1 To Class(i).startSpellCount
                Class(i).StartSpell(X) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & X))
                Class(i).StartSpellLevel(X) = Val(GetVar(filename, "CLASS" & i, "StartSpellLevel" & X))
            Next
        End If
    Next
End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim X As Long
    
    filename = App.path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        
        ' loop for items & values
        For X = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & X, STR(Class(i).StartItem(X)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & X, STR(Class(i).StartValue(X)))
        Next
        
        ' loop for spells
        For X = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & X, STR(Class(i).StartSpell(X)))
            Call PutVar(filename, "CLASS" & i, "StartSpellLevel" & X, STR(Class(i).StartSpellLevel(X)))
        Next
    Next
End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If
End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next
End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemnum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.path & "\data\items\item" & itemnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(itemnum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\spells\spells" & spellNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(spellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal npcNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\npcs\npc" & npcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , NPC(npcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , NPC(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim X As Long
    Dim Y As Long, i As Long, z As Long, w As Long
    filename = App.path & "\data\maps\map" & mapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(mapNum).Name
    Put #F, , Map(mapNum).Music
    Put #F, , Map(mapNum).BGS
    Put #F, , Map(mapNum).Revision
    Put #F, , Map(mapNum).Moral
    Put #F, , Map(mapNum).Up
    Put #F, , Map(mapNum).Down
    Put #F, , Map(mapNum).Left
    Put #F, , Map(mapNum).Right
    Put #F, , Map(mapNum).BootMap
    Put #F, , Map(mapNum).BootX
    Put #F, , Map(mapNum).BootY
    
    Put #F, , Map(mapNum).Weather
    Put #F, , Map(mapNum).WeatherIntensity
    
    Put #F, , Map(mapNum).Fog
    Put #F, , Map(mapNum).FogSpeed
    Put #F, , Map(mapNum).FogOpacity
    
    Put #F, , Map(mapNum).Red
    Put #F, , Map(mapNum).Green
    Put #F, , Map(mapNum).Blue
    Put #F, , Map(mapNum).Alpha
    
    Put #F, , Map(mapNum).MaxX
    Put #F, , Map(mapNum).MaxY

    For X = 0 To Map(mapNum).MaxX
        For Y = 0 To Map(mapNum).MaxY
            Put #F, , Map(mapNum).Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , Map(mapNum).NPC(X)
        Put #F, , Map(mapNum).NpcSpawnType(X)
    Next
    Close #F
    
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    filename = App.path & "\data\maps\map" & mapNum & "_eventdata.dat"
    PutVar filename, "Events", "EventCount", Val(Map(mapNum).EventCount)
    
    If Map(mapNum).EventCount > 0 Then
        For i = 1 To Map(mapNum).EventCount
            With Map(mapNum).Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "Global", Val(.Global)
                PutVar filename, "Event" & i, "x", Val(.X)
                PutVar filename, "Event" & i, "y", Val(.Y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(mapNum).Events(i).PageCount > 0 Then
                For X = 1 To Map(mapNum).Events(i).PageCount
                    With Map(mapNum).Events(i).Pages(X)
                        PutVar filename, "Event" & i & "Page" & X, "chkVariable", Val(.chkVariable)
                        PutVar filename, "Event" & i & "Page" & X, "VariableIndex", Val(.VariableIndex)
                        PutVar filename, "Event" & i & "Page" & X, "VariableCondition", Val(.VariableCondition)
                        PutVar filename, "Event" & i & "Page" & X, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar filename, "Event" & i & "Page" & X, "chkSwitch", Val(.chkSwitch)
                        PutVar filename, "Event" & i & "Page" & X, "SwitchIndex", Val(.SwitchIndex)
                        PutVar filename, "Event" & i & "Page" & X, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & X, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & X, "HasItemIndex", Val(.HasItemIndex)
                        
                        PutVar filename, "Event" & i & "Page" & X, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & X, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar filename, "Event" & i & "Page" & X, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & X, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & X, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicX2", Val(.GraphicX2)
                        PutVar filename, "Event" & i & "Page" & X, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar filename, "Event" & i & "Page" & X, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & X, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & X, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar filename, "Event" & i & "Page" & X, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar filename, "Event" & i & "Page" & X, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar filename, "Event" & i & "Page" & X, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For Y = 1 To .MoveRouteCount
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Index", Val(.MoveRoute(Y).Index)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data1", Val(.MoveRoute(Y).Data1)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data2", Val(.MoveRoute(Y).Data2)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data3", Val(.MoveRoute(Y).Data3)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data4", Val(.MoveRoute(Y).Data4)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data5", Val(.MoveRoute(Y).data5)
                                PutVar filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data6", Val(.MoveRoute(Y).data6)
                            Next
                        End If
                        
                        PutVar filename, "Event" & i & "Page" & X, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & X, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & X, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & X, "ShowName", Val(.ShowName)
                        PutVar filename, "Event" & i & "Page" & X, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & X, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar filename, "Event" & i & "Page" & X, "Position", Val(.Position)
                    End With
                    
                    If Map(mapNum).Events(i).Pages(X).CommandListCount > 0 Then
                        For Y = 1 To Map(mapNum).Events(i).Pages(X).CommandListCount
                            PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "CommandCount", Val(Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount)
                            PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "ParentList", Val(Map(mapNum).Events(i).Pages(X).CommandList(Y).ParentList)
                            If Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                For z = 1 To Map(mapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map(mapNum).Events(i).Pages(X).CommandList(Y).Commands(z)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Index", Val(.Index)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text1", .Text1
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text2", .Text2
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text3", .Text3
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text4", .Text4
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Text5", .Text5
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data5", Val(.data5)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "Data6", Val(.data6)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).Index)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).data5)
                                                PutVar filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).data6)
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
    
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim X As Long
    Dim Y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).BGS
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        
        Get #F, , Map(i).Weather
        Get #F, , Map(i).WeatherIntensity
        
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For X = 0 To Map(i).MaxX
            For Y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(X, Y)
            Next
        Next

        For X = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).NPC(X)
            Get #F, , Map(i).NpcSpawnType(X)
            MapNpc(i).NPC(X).Num = Map(i).NPC(X)
        Next

        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        filename = App.path & "\data\maps\map" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(filename, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                With Map(z).Events(i)
                    .Name = GetVar(filename, "Event" & i, "Name")
                    .Global = Val(GetVar(filename, "Event" & i, "Global"))
                    .X = Val(GetVar(filename, "Event" & i, "x"))
                    .Y = Val(GetVar(filename, "Event" & i, "y"))
                    .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
                End With
                If Map(z).Events(i).PageCount > 0 Then
                    ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                    For X = 1 To Map(z).Events(i).PageCount
                        With Map(z).Events(i).Pages(X)
                            .chkVariable = Val(GetVar(filename, "Event" & i & "Page" & X, "chkVariable"))
                            .VariableIndex = Val(GetVar(filename, "Event" & i & "Page" & X, "VariableIndex"))
                            .VariableCondition = Val(GetVar(filename, "Event" & i & "Page" & X, "VariableCondition"))
                            .VariableCompare = Val(GetVar(filename, "Event" & i & "Page" & X, "VariableCompare"))
                            
                            .chkSwitch = Val(GetVar(filename, "Event" & i & "Page" & X, "chkSwitch"))
                            .SwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & X, "SwitchIndex"))
                            .SwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & X, "SwitchCompare"))
                            
                            .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & X, "chkHasItem"))
                            .HasItemIndex = Val(GetVar(filename, "Event" & i & "Page" & X, "HasItemIndex"))
                            
                            .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & X, "chkSelfSwitch"))
                            .SelfSwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & X, "SelfSwitchIndex"))
                            .SelfSwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & X, "SelfSwitchCompare"))
                            
                            .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicType"))
                            .Graphic = Val(GetVar(filename, "Event" & i & "Page" & X, "Graphic"))
                            .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicX"))
                            .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicY"))
                            .GraphicX2 = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicX2"))
                            .GraphicY2 = Val(GetVar(filename, "Event" & i & "Page" & X, "GraphicY2"))
                            
                            .MoveType = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveType"))
                            .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveSpeed"))
                            .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveFreq"))
                            
                            .IgnoreMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & X, "IgnoreMoveRoute"))
                            .RepeatMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & X, "RepeatMoveRoute"))
                            
                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRouteCount"))
                            
                            If .MoveRouteCount > 0 Then
                                ReDim Map(z).Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                                For Y = 1 To .MoveRouteCount
                                    .MoveRoute(Y).Index = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Index"))
                                    .MoveRoute(Y).Data1 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data1"))
                                    .MoveRoute(Y).Data2 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data2"))
                                    .MoveRoute(Y).Data3 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data3"))
                                    .MoveRoute(Y).Data4 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data4"))
                                    .MoveRoute(Y).data5 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data5"))
                                    .MoveRoute(Y).data6 = Val(GetVar(filename, "Event" & i & "Page" & X, "MoveRoute" & Y & "Data6"))
                                Next
                            End If
                            
                            .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & X, "WalkAnim"))
                            .DirFix = Val(GetVar(filename, "Event" & i & "Page" & X, "DirFix"))
                            .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & X, "WalkThrough"))
                            .ShowName = Val(GetVar(filename, "Event" & i & "Page" & X, "ShowName"))
                            .Trigger = Val(GetVar(filename, "Event" & i & "Page" & X, "Trigger"))
                            .CommandListCount = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandListCount"))
                         
                            .Position = Val(GetVar(filename, "Event" & i & "Page" & X, "Position"))
                        End With
                            
                        If Map(z).Events(i).Pages(X).CommandListCount > 0 Then
                            ReDim Map(z).Events(i).Pages(X).CommandList(0 To Map(z).Events(i).Pages(X).CommandListCount)
                            For Y = 1 To Map(z).Events(i).Pages(X).CommandListCount
                                Map(z).Events(i).Pages(X).CommandList(Y).CommandCount = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "CommandCount"))
                                Map(z).Events(i).Pages(X).CommandList(Y).ParentList = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "ParentList"))
                                If Map(z).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(X).CommandList(Y).Commands(Map(z).Events(i).Pages(X).CommandList(Y).CommandCount)
                                    For p = 1 To Map(z).Events(i).Pages(X).CommandList(Y).CommandCount
                                        With Map(z).Events(i).Pages(X).CommandList(Y).Commands(p)
                                            .Index = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Index"))
                                            .Text1 = GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text1")
                                            .Text2 = GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text2")
                                            .Text3 = GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text3")
                                            .Text4 = GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text4")
                                            .Text5 = GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Text5")
                                            .Data1 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data1"))
                                            .Data2 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data2"))
                                            .Data3 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data3"))
                                            .Data4 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data4"))
                                            .data5 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data5"))
                                            .data6 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "Data6"))
                                            .ConditionalBranch.CommandList = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchCommandList"))
                                            .ConditionalBranch.Condition = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchCondition"))
                                            .ConditionalBranch.Data1 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData1"))
                                            .ConditionalBranch.Data2 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData2"))
                                            .ConditionalBranch.Data3 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchData3"))
                                            .ConditionalBranch.ElseCommandList = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "ConditionalBranchElseCommandList"))
                                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRouteCount"))
                                            If .MoveRouteCount > 0 Then
                                                ReDim .MoveRoute(1 To .MoveRouteCount)
                                                For w = 1 To .MoveRouteCount
                                                    .MoveRoute(w).Index = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Index"))
                                                    .MoveRoute(w).Data1 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                    .MoveRoute(w).Data2 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                    .MoveRoute(w).Data3 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                    .MoveRoute(w).Data4 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                    .MoveRoute(w).data5 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                    .MoveRoute(w).data6 = Val(GetVar(filename, "Event" & i & "Page" & X, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data6"))
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
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal mapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapNum, Index)), LenB(MapItem(mapNum, Index)))
    MapItem(mapNum, Index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal mapNum As Long)
    ReDim MapNpc(mapNum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapNum).NPC(Index)), LenB(MapNpc(mapNum).NPC(Index)))
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next

End Sub

Sub ClearMap(ByVal mapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(mapNum)), LenB(Map(mapNum)))
    Map(mapNum).Name = vbNullString
    Map(mapNum).MaxX = MAX_MAPX
    Map(mapNum).MaxY = MAX_MAPY
    ReDim Map(mapNum).Tile(0 To Map(mapNum).MaxX, 0 To Map(mapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    Call ClearBank(Index)

    filename = App.path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F

End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub ClearParty(ByVal partyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String
filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

End Sub

Sub SaveVariables()
Dim i As Long, filename As String
filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
filename = App.path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
filename = App.path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
Next
End Sub
