VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraServer 
      Caption         =   "Servidor"
      Height          =   3495
      Left            =   5880
      TabIndex        =   16
      Top             =   3240
      Width           =   5175
      Begin VB.CommandButton cmdShutDown 
         Caption         =   "Desligar"
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         ToolTipText     =   "Fechar servidor com aviso previo."
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Fechar"
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         ToolTipText     =   "Desligar servidor sem aviso."
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CheckBox chkServerLog 
         Caption         =   "Server Log"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Habilitar ou Desabilitar log de informações"
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Recarregar"
      Height          =   3495
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
      Begin VB.CommandButton cmbReloadQuests 
         Caption         =   "Missões"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadClasses 
         Caption         =   "Classes"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadMaps 
         Caption         =   "Mapas"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CmdReloadSpells 
         Caption         =   "Magias"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadShops 
         Caption         =   "Lojas"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadNPCs 
         Caption         =   "NPCs"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadItems 
         Caption         =   "Items"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdReloadResources 
         Caption         =   "Recursos"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame fraPlayers 
      Caption         =   "Jogadores"
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4095
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5530
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Endereço IP"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Conta"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Personagem"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Frame fraConsole 
      Caption         =   "Console"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.TextBox txtChat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   2610
         Width           =   10695
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   10695
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8400
         TabIndex        =   4
         ToolTipText     =   "Travar ou Destravar CPS"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblCPS 
         AutoSize        =   -1  'True
         Caption         =   "CPS: 0"
         Height          =   195
         Left            =   9240
         TabIndex        =   3
         ToolTipText     =   "Ciclos por segundo"
         Top             =   0
         Width           =   600
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Player(1).Switches(1) = 1
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmbReloadQuests_Click()
    Dim i As Integer
    
    Call LoadQuests ' Carregar Missões
    
    Call TextAdd("Todas as missões recarregadas!")
    
    ' Enviar classes para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SendQuests(i)
        End If
    Next
End Sub

Private Sub cmdReloadClasses_Click()
    Dim i As Integer
    
    Call LoadClasses ' Carregar Classes
    
    Call TextAdd("Todas as classe recarregadas!")
    
    ' Enviar classes para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
    Dim i As Integer
    
    Call LoadItems ' Carregar itens
    
    Call TextAdd("Todos os itens recarregados!")
    
    ' Enviar itens para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
    Dim i As Integer
    
    Call LoadMaps ' Carregar Mapas
    
    Call TextAdd("Todos os mapas recarregados!")
    
    ' Enviar mapa para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
    Dim i As Integer
    
    Call LoadNpcs ' Carregar NPCs
    
    Call TextAdd("Todos os NPCs recarregados!")
    
    ' Enviar NPCs para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
    Dim i As Integer
    
    Call LoadShops ' Carregar lojas
    
    Call TextAdd("Todas as lojas recarregadas!")
    
    ' Enviar lojas para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
    Dim i As Integer
    
    Call LoadSpells ' Carregar magias
    
    Call TextAdd("Todas as magias recarregadas!")
    
    ' Enviar magias para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
    Dim i As Integer
    
    Call LoadResources ' Carregar Recursos
    
    Call TextAdd("Todos os recursos recarregados!")
    
    ' Enviar para todos
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Desligar"
        GlobalMsg "Desligamento cancelado!", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancelar"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lmsg As Long
    lmsg = X / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub
