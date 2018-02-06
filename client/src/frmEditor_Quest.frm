VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest System"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   16875
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Titulo da Missão:"
      Height          =   975
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Recompensas"
         Height          =   300
         Index           =   2
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Objetivos"
         Height          =   300
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Requerimentos"
         Height          =   300
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Dados"
         Height          =   300
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Deletar"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salvar"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lista de Missões"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdArray 
         Caption         =   "Alterar Tamanho da Lista"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   7200
         Width           =   3135
      End
      Begin VB.ListBox lstIndex 
         Appearance      =   0  'Flat
         Height          =   6855
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTasks 
      Caption         =   "Objetivos"
      Height          =   6495
      Left            =   3600
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   12975
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   6960
         TabIndex        =   87
         Top             =   480
         Width           =   2775
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   89
            Top             =   2880
            Width           =   2535
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   88
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2640
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            Caption         =   "Map: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   91
            Top             =   2640
            Width           =   525
         End
         Begin VB.Label lblResource 
            AutoSize        =   -1  'True
            Caption         =   "Resource: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   90
            Top             =   3240
            Width           =   915
         End
      End
      Begin VB.Frame fraTask 
         Caption         =   "Objetivo: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   6135
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   5055
         Begin VB.Frame fraTaskItem 
            Caption         =   "Item: Nenhum"
            ForeColor       =   &H80000011&
            Height          =   615
            Left            =   2280
            TabIndex        =   96
            Top             =   1810
            Visible         =   0   'False
            Width           =   2655
            Begin VB.HScrollBar scrlItem 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   97
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame fraTaskNPC 
            Caption         =   "NPC: Nenhum"
            ForeColor       =   &H80000011&
            Height          =   615
            Left            =   2280
            TabIndex        =   94
            Top             =   1810
            Visible         =   0   'False
            Width           =   2655
            Begin VB.HScrollBar scrlNPC 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   95
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame fraAmount 
            Caption         =   "Quantidade: 1"
            ForeColor       =   &H80000011&
            Height          =   615
            Left            =   2280
            TabIndex        =   92
            Top             =   5400
            Visible         =   0   'False
            Width           =   2655
            Begin VB.HScrollBar scrlAmount 
               Height          =   255
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   93
               Top             =   240
               Value           =   1
               Width           =   2415
            End
         End
         Begin VB.Frame Frame1 
            Height          =   4200
            Left            =   120
            TabIndex        =   77
            Top             =   1810
            Width           =   2055
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Obter de NPC"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   8
               Left            =   120
               TabIndex        =   86
               Top             =   2280
               Width           =   1815
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Pegar Recurso"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   7
               Left            =   120
               TabIndex        =   85
               Top             =   2040
               Width           =   1815
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Matar Jogador"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   6
               Left            =   120
               TabIndex        =   84
               Top             =   1800
               Width           =   1695
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Dar item a um NPC"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   5
               Left            =   120
               TabIndex        =   83
               Top             =   1560
               Width           =   1815
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Ir para um local"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   4
               Left            =   120
               TabIndex        =   82
               Top             =   1320
               Width           =   1695
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Falar com NPC"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   120
               TabIndex        =   81
               Top             =   1080
               Width           =   1695
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Obter Item"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   80
               Top             =   840
               Width           =   1695
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Mata NPC"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   79
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optTask 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "Não faz nada"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000000&
               X1              =   120
               X2              =   2040
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.CheckBox chkEnd 
            Caption         =   "Objetivo Final?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Left            =   120
            TabIndex        =   75
            Top             =   600
            Width           =   1455
         End
         Begin VB.HScrollBar scrlTotalTasks 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   72
            Top             =   240
            Value           =   1
            Width           =   4815
         End
         Begin VB.Frame fraTaskText 
            Caption         =   "Mensagem:"
            ForeColor       =   &H80000011&
            Height          =   975
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   4815
            Begin VB.TextBox txtTaskLog 
               Appearance      =   0  'Flat
               Height          =   615
               Left            =   120
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Top             =   240
               Width           =   4575
            End
         End
         Begin VB.Label lblInfo 
            Caption         =   "Marque esse opção caso deseje que a missão acabe junto a este objetivo!"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   76
            Top             =   545
            Width           =   3255
         End
      End
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Recompensas"
      Height          =   6495
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame fraItensReward 
         Caption         =   "Itens"
         ForeColor       =   &H80000011&
         Height          =   4215
         Left            =   120
         TabIndex        =   64
         Top             =   2080
         Width           =   5055
         Begin VB.CommandButton cmdItemRewRemove 
            Caption         =   "X"
            Height          =   255
            Left            =   4440
            TabIndex        =   70
            Top             =   3840
            Width           =   495
         End
         Begin VB.CommandButton cmdItemRew 
            Caption         =   "Atualizar"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   3840
            Width           =   2415
         End
         Begin VB.ListBox lstItemRew 
            Appearance      =   0  'Flat
            Height          =   2550
            ItemData        =   "frmEditor_Quest.frx":0000
            Left            =   120
            List            =   "frmEditor_Quest.frx":0007
            TabIndex        =   68
            Top             =   1190
            Width           =   4800
         End
         Begin VB.Frame fraRewardItem 
            Caption         =   "Item: Nenhum x1"
            ForeColor       =   &H80000010&
            Height          =   975
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   4815
            Begin VB.HScrollBar scrlItemRewValue 
               Height          =   255
               Left            =   120
               Max             =   10
               Min             =   1
               TabIndex        =   67
               Top             =   600
               Value           =   1
               Width           =   4575
            End
            Begin VB.HScrollBar scrlItemRew 
               Height          =   255
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   66
               Top             =   240
               Value           =   1
               Width           =   4575
            End
         End
      End
      Begin VB.Frame fraRewardTitle 
         Caption         =   "Titulo: Nenhum"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   120
         TabIndex        =   62
         Top             =   1470
         Width           =   5055
         Begin VB.HScrollBar scrlRewardTitle 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraRewardMoney 
         Caption         =   "Dinheiro: 0"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   120
         TabIndex        =   60
         Top             =   855
         Width           =   5055
         Begin VB.HScrollBar scrlRewardMoney 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraRewardEXP 
         Caption         =   "Experiência: 0"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   5055
         Begin VB.HScrollBar scrlExp 
            Height          =   255
            LargeChange     =   50
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requerimentos"
      Height          =   6495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame fraReqItens 
         Caption         =   "Itens"
         ForeColor       =   &H80000011&
         Height          =   2730
         Left            =   120
         TabIndex        =   51
         Top             =   3640
         Width           =   5055
         Begin VB.CommandButton cmdReqItemRemove 
            Caption         =   "X"
            Height          =   255
            Left            =   4440
            TabIndex        =   57
            Top             =   2350
            Width           =   495
         End
         Begin VB.CommandButton cmdReqItem 
            Caption         =   "Atualizar"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   2350
            Width           =   2295
         End
         Begin VB.ListBox lstReqItem 
            Appearance      =   0  'Flat
            Height          =   1110
            ItemData        =   "frmEditor_Quest.frx":0017
            Left            =   120
            List            =   "frmEditor_Quest.frx":0019
            TabIndex        =   55
            Top             =   1190
            Width           =   4805
         End
         Begin VB.Frame fraReqItem 
            Caption         =   "Item: Nenhum x1"
            ForeColor       =   &H80000010&
            Height          =   975
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   4815
            Begin VB.HScrollBar scrlReqItemValue 
               Height          =   255
               Left            =   120
               Max             =   10
               Min             =   1
               TabIndex        =   54
               Top             =   600
               Value           =   1
               Width           =   4575
            End
            Begin VB.HScrollBar scrlReqItem 
               Height          =   255
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   53
               Top             =   240
               Value           =   1
               Width           =   4575
            End
         End
      End
      Begin VB.Frame fraReqClass 
         Caption         =   "Classes"
         ForeColor       =   &H80000011&
         Height          =   2175
         Left            =   120
         TabIndex        =   45
         Top             =   1470
         Width           =   5055
         Begin VB.CommandButton cmdReqClassRemove 
            Caption         =   "X"
            Height          =   255
            Left            =   4440
            TabIndex        =   50
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdReqClass 
            Caption         =   "Atualizar"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1800
            Width           =   2295
         End
         Begin VB.ListBox lstReqClass 
            Appearance      =   0  'Flat
            Height          =   930
            ItemData        =   "frmEditor_Quest.frx":001B
            Left            =   120
            List            =   "frmEditor_Quest.frx":001D
            TabIndex        =   48
            Top             =   830
            Width           =   4800
         End
         Begin VB.Frame fraReqClassId 
            Caption         =   "Classe: Nenhuma"
            ForeColor       =   &H80000010&
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   4815
            Begin VB.HScrollBar scrlReqClass 
               Height          =   255
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   47
               Top             =   240
               Value           =   1
               Width           =   4575
            End
         End
      End
      Begin VB.Frame fraReqQuest 
         Caption         =   "Missão: Nenhuma"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   855
         Width           =   5055
         Begin VB.HScrollBar scrlReqQuest 
            Height          =   255
            Left            =   120
            Max             =   70
            TabIndex        =   44
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraReqLevel 
         Caption         =   "Level: 0"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   5055
         Begin VB.HScrollBar scrlReqLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   42
            Top             =   240
            Width           =   4815
         End
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "Informações Gerais"
      Height          =   6495
      Left            =   3600
      TabIndex        =   11
      Top             =   1200
      Width           =   5295
      Begin VB.Frame fraItens 
         Caption         =   "Itens:"
         ForeColor       =   &H80000011&
         Height          =   2175
         Left            =   240
         TabIndex        =   31
         Top             =   4050
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlCItemValue 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   1
            TabIndex        =   37
            Top             =   960
            Value           =   1
            Width           =   2775
         End
         Begin VB.PictureBox picCIcon 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   4180
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   35
            Top             =   330
            Width           =   510
         End
         Begin VB.Frame fraCItem 
            Caption         =   "Item: Nenhuma"
            ForeColor       =   &H80000011&
            Height          =   620
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3975
            Begin VB.HScrollBar scrlCItemId 
               Height          =   255
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   36
               Top             =   240
               Value           =   1
               Width           =   3735
            End
         End
         Begin VB.CommandButton cmdCancelC 
            Caption         =   "Cancelar"
            Height          =   300
            Left            =   2640
            TabIndex        =   33
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton cmdConfirmC 
            Caption         =   "Confirmar"
            Height          =   300
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblCValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblTValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Quantidade:"
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame fraComerce 
         Caption         =   "Comércio:"
         ForeColor       =   &H80000011&
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         Top             =   3810
         Width           =   5055
         Begin VB.Frame Frame4 
            Caption         =   "Tomar itens no final:"
            ForeColor       =   &H80000010&
            Height          =   2175
            Left            =   2640
            TabIndex        =   27
            Top             =   240
            Width           =   2295
            Begin VB.CommandButton cmdTakeItemRemove 
               Caption         =   "X"
               Height          =   255
               Left            =   1800
               TabIndex        =   28
               Top             =   1800
               Width           =   375
            End
            Begin VB.CommandButton cmdTakeItem 
               Caption         =   "Atualizar"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   1800
               Width           =   1575
            End
            Begin VB.ListBox lstTakeItem 
               Appearance      =   0  'Flat
               Height          =   1575
               IntegralHeight  =   0   'False
               ItemData        =   "frmEditor_Quest.frx":001F
               Left            =   120
               List            =   "frmEditor_Quest.frx":0026
               TabIndex        =   30
               Top             =   240
               Width           =   2055
            End
         End
         Begin VB.Frame fraStartItem 
            Caption         =   "Doar itens no inicio:"
            ForeColor       =   &H80000010&
            Height          =   2175
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2295
            Begin VB.CommandButton cmdGiveItemRemove 
               Caption         =   "X"
               Height          =   255
               Left            =   1800
               TabIndex        =   26
               Top             =   1800
               Width           =   375
            End
            Begin VB.CommandButton cmdGiveItem 
               Caption         =   "Atualizar"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1800
               Width           =   1575
            End
            Begin VB.ListBox lstGiveItem 
               Appearance      =   0  'Flat
               Height          =   1575
               IntegralHeight  =   0   'False
               ItemData        =   "frmEditor_Quest.frx":0037
               Left            =   120
               List            =   "frmEditor_Quest.frx":003E
               TabIndex        =   24
               Top             =   240
               Width           =   2055
            End
         End
      End
      Begin VB.Frame fraFinish 
         Caption         =   "Mensagem ao concluir:"
         ForeColor       =   &H80000011&
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   2710
         Width           =   5055
         Begin VB.TextBox txtFinishMessage 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraMidText 
         Caption         =   "Mensagem do NPC caso não tenha sido concluida:"
         ForeColor       =   &H80000011&
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   1615
         Width           =   5055
         Begin VB.TextBox txtMidMessage 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.CheckBox chkRepeat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Missão Repetitiva?"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3340
         MaskColor       =   &H00404040&
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame fraStartText 
         Caption         =   "História:"
         ForeColor       =   &H80000011&
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   405
         Width           =   5055
         Begin VB.TextBox txtStartMessage 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   4815
         End
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////

Option Explicit

Private TempTask As Long
Private GiveOrTake As Byte ' Variável que irá informar tipo

Private Sub Form_Load()
    scrlTotalTasks.max = MAX_TASKS
    scrlNPC.max = MAX_NPCS
    scrlItem.max = MAX_ITEMS
    scrlMap.max = MAX_MAPS
    scrlResource.max = MAX_RESOURCES
    scrlAmount.max = MAX_BYTE
    scrlReqLevel.max = MAX_LEVELS
    scrlReqQuest.max = MAX_QUESTS
    scrlReqClass.max = Max_Classes
    
    scrlReqItem.max = MAX_ITEMS
    scrlReqItemValue.max = MAX_BYTE
    
    scrlCItemId.max = MAX_ITEMS
    scrlCItemValue.max = MAX_BYTE
    scrlExp.max = MAX_INTEGER 'Alatar v1.2
    scrlItemRew.max = MAX_ITEMS
    scrlItemRewValue.max = MAX_BYTE
    
    
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
End Sub

Private Sub lstIndex_Click()
    QuestEditorInit
End Sub

Private Sub optTask_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Order = Index
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtTaskLog_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskLog = Trim$(txtTaskLog.Text)
End Sub

Private Sub chkRepeat_Click()
    If chkRepeat.Value = 1 Then
        Quest(EditorIndex).Repeat = 1
    Else
        Quest(EditorIndex).Repeat = 0
    End If
End Sub

'Alatar v1.2

Private Sub cmdReqItem_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlReqItem.Value < 1 Or scrlReqItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlReqItem.Value).Name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = scrlReqItem.Value
    Quest(EditorIndex).RequiredItem(Index).Value = scrlReqItemValue.Value
    UpdateQuestRequirementItems
End Sub

Private Sub cmdReqItemRemove_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = 0
    Quest(EditorIndex).RequiredItem(Index).Value = 1
    UpdateQuestRequirementItems
End Sub

'/Alatar v1.2

'Alatar v1.2
Private Sub cmdItemRew_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlItemRew.Value < 1 Or scrlItemRew.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlItemRew.Value).Name) = "" Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = scrlItemRew.Value
    Quest(EditorIndex).RewardItem(Index).Value = scrlItemRewValue.Value
    UpdateQuestRewardItems
End Sub

Private Sub cmdItemRewRemove_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = 0
    Quest(EditorIndex).RewardItem(Index).Value = 1
    UpdateQuestRewardItems
End Sub
'/Alatar v1.2

Private Sub scrlMap_Change()
    lblMap.Caption = "Map: " & scrlMap.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Map = scrlMap.Value
End Sub

Private Sub scrlResource_Change()
    lblResource.Caption = "Resource: " & scrlResource.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Resource = scrlResource.Value
End Sub

Private Sub chkEnd_Click()
    If chkEnd.Value = 1 Then
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = True
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = False
    End If
End Sub

Private Sub optShowFrame_Click(Index As Integer)
    fraGeneral.Visible = False
    fraRequirements.Visible = False
    fraRewards.Visible = False
    fraTasks.Visible = False
    
    If optShowFrame(Index).Value = True Then
        Select Case Index
            Case 0
                fraGeneral.Visible = True
            Case 1
                fraRequirements.Visible = True
            Case 2
                fraRewards.Visible = True
            Case 3
                fraTasks.Visible = True
        End Select
    End If
End Sub

' DOOOLLY MODIFIER

' Informações Gerais - ###############################################

Private Sub scrlCItemId_Change()
    ' Altera texto
    fraCItem.Caption = "Item: " & Trim(Item(scrlCItemId.Value).Name)
    
    ' Atualizar barra de quantidade
    If Item(scrlCItemId.Value).Type <> ITEM_TYPE_CURRENCY Then
        scrlCItemValue.Value = 1
        scrlCItemValue.Enabled = False
    Else
        scrlCItemValue.Enabled = True
    End If
End Sub

Private Sub scrlCItemValue_Change()
    ' Alterar texto
    lblCValue.Caption = scrlCItemValue.Value
End Sub

Private Sub cmdGiveItem_Click()
    ' Atualizar visibilidade
    fraItens.Visible = True
    
    ' Resetar scrolls
    scrlCItemId.Value = 1
    scrlCItemValue.Value = 1
    
    GiveOrTake = 0 ' Dar item
End Sub

Private Sub cmdTakeItem_Click()
    ' Atualizar visibilidade
    fraItens.Visible = True
    
    ' Resetar scrolls
    scrlCItemId.Value = 1
    scrlCItemValue.Value = 1
    
    GiveOrTake = 1 ' Tomar item
End Sub

Private Sub lstGiveItem_Click()
    If lstGiveItem.ListIndex >= 0 Then
        cmdGiveItem.Enabled = True
    End If
End Sub

Private Sub lstTakeItem_Click()
    If lstTakeItem.ListIndex >= 0 Then
        cmdTakeItem.Enabled = True
    End If
End Sub

Private Sub cmdCancelC_Click()
    fraItens.Visible = False
End Sub

Private Sub cmdGiveItemRemove_Click()
    Dim Index As Long
    
    Index = lstGiveItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).GiveItem(Index).Item = 0
    Quest(EditorIndex).GiveItem(Index).Value = 1
    Call UpdateQuestGiveItems
End Sub

Private Sub cmdTakeItemRemove_Click()
    Dim Index As Long
    
    Index = lstTakeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).TakeItem(Index).Item = 0
    Quest(EditorIndex).TakeItem(Index).Value = 1
    UpdateQuestTakeItems
End Sub

Private Sub txtFinishMessage_Change()
    Quest(EditorIndex).FinishMessage = Trim$(txtFinishMessage.Text)
End Sub

Private Sub txtMidMessage_Change()
    Quest(EditorIndex).MidMessage = Trim$(txtMidMessage.Text)
End Sub

Private Sub txtStartMessage_Change()
    Quest(EditorIndex).StartMessage = Trim$(txtStartMessage.Text)
End Sub

Private Sub cmdConfirmC_Click()
    Dim Index As Long
    
    Select Case GiveOrTake
        Case 0 ' Dar item
            Index = lstGiveItem.ListIndex + 1 'the selected item
            If Index = 0 Or scrlCItemId.Value < 1 Or scrlCItemId.Value > MAX_ITEMS Then Exit Sub ' Tratamento de erros
            If Trim$(Item(scrlCItemId.Value).Name) = "" Then Exit Sub
            
            Quest(EditorIndex).GiveItem(Index).Item = scrlCItemId.Value
            Quest(EditorIndex).GiveItem(Index).Value = scrlCItemValue.Value
            
            Call UpdateQuestGiveItems
        Case 1 ' Pegar item
            Index = lstTakeItem.ListIndex + 1 'the selected item
            If Index = 0 Or scrlCItemId.Value < 1 Or scrlCItemId.Value > MAX_ITEMS Then Exit Sub ' Tratamento de erros
            If Trim$(Item(scrlCItemId.Value).Name) = "" Then Exit Sub
    
            Quest(EditorIndex).TakeItem(Index).Item = scrlCItemId.Value
            Quest(EditorIndex).TakeItem(Index).Value = scrlCItemValue.Value
            
            Call UpdateQuestTakeItems
    End Select
    
    fraItens.Visible = False
End Sub

' Requerimentos ############################################################

Private Sub scrlReqLevel_Change()
    fraReqLevel.Caption = "Level: " & scrlReqLevel.Value
    Quest(EditorIndex).RequiredLevel = scrlReqLevel.Value
End Sub

Private Sub scrlReqQuest_Change()
    If Not scrlReqQuest.Value = 0 Then
        If Not Trim$(Quest(scrlReqQuest.Value).Name) = "" Then
            fraReqQuest.Caption = "Missão: " & Trim$(Quest(scrlReqQuest.Value).Name)
        Else
            fraReqQuest.Caption = "Missão: (Sem nome)ID: " & scrlReqQuest.Value
        End If
    Else
        fraReqQuest.Caption = "Missão: Nenhuma"
    End If
    
    Quest(EditorIndex).RequiredQuest = scrlReqQuest.Value
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then
        fraReqClassId.Caption = "Classe: Nenhuma"
    Else
        fraReqClassId.Caption = "Classe: " & Trim$(Class(scrlReqClass.Value).Name)
    End If
End Sub

Private Sub cmdReqClass_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1 'the selected class
    If Index = 0 Then Exit Sub
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then Exit Sub
    If Trim$(Class(scrlReqClass.Value).Name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = scrlReqClass.Value
    UpdateQuestClass
End Sub

Private Sub cmdReqClassRemove_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = 0
    UpdateQuestClass
End Sub

Private Sub scrlReqItem_Change()
    If Not scrlReqItem.Value = 0 Then
        If Not Trim(Item(scrlReqItem.Value).Name) = vbNullString Then
            fraReqItem.Caption = "Item: " & Trim(Item(scrlReqItem.Value).Name) & " x" & scrlReqItemValue.Value
        Else
            fraReqItem.Caption = "Item: (Sem nome)ID: " & scrlReqItem.Value & " x" & scrlReqItemValue.Value
        End If
    Else
        fraReqItem.Caption = "Item: Nenhum"
    End If
End Sub

Private Sub scrlReqItemValue_Change()
    If Not scrlReqItem.Value = 0 Then
        If Not Trim(Item(scrlReqItem.Value).Name) = vbNullString Then
            fraReqItem.Caption = "Item: " & Trim$(Item(scrlReqItem.Value).Name) & " x" & scrlReqItemValue.Value
        Else
            fraReqItem.Caption = "Item: (Sem nome)ID: " & scrlReqItem.Value & " x" & scrlReqItemValue.Value
        End If
    Else
        fraReqItem.Caption = "Item: Nenhum"
    End If
End Sub

' Recompensas - ###################################################

Private Sub scrlExp_Change()
    fraRewardEXP.Caption = "Experiência: " & scrlExp.Value
    Quest(EditorIndex).RewardExp = scrlExp.Value
End Sub

Private Sub scrlRewardMoney_Change()
    fraRewardMoney.Caption = "Dinheiro: " & scrlRewardMoney.Value
    Quest(EditorIndex).RewardMoney = scrlRewardMoney.Value
End Sub

Private Sub scrlRewardTitle_Change()
    fraRewardTitle.Caption = "Titulo: " & scrlRewardTitle.Value
    Quest(EditorIndex).RewardTitle = scrlRewardTitle.Value
End Sub

Private Sub scrlItemRew_Change()
    If Not scrlItemRew.Value = 0 Then
        If Not Trim(Item(scrlItemRew.Value).Name) = vbNullString Then
            fraRewardItem.Caption = "Item: " & Trim$(Item(scrlItemRew.Value).Name) & " x" & scrlItemRewValue.Value
        Else
            fraRewardItem.Caption = "Item: (Sem nome)ID: " & scrlItemRew.Value & " x" & scrlItemRewValue.Value
        End If
    Else
        fraRewardItem.Caption = "Item: Nenhum"
    End If
End Sub

Private Sub scrlItemRewValue_Change()
    If Not scrlItemRew.Value = 0 Then
        If Not Trim(Item(scrlItemRew.Value).Name) = vbNullString Then
            fraRewardItem.Caption = "Item: " & Trim$(Item(scrlItemRew.Value).Name) & " x" & scrlItemRewValue.Value
        Else
            fraRewardItem.Caption = "Item: (Sem nome)ID: " & scrlItemRew.Value & " x" & scrlItemRewValue.Value
        End If
    Else
        fraRewardItem.Caption = "Item: Nenhum"
    End If
End Sub

' Objetivos - ###################################################

Private Sub scrlTotalTasks_Change()
    Dim i As Long
    
    fraTask.Caption = "Objetivo: " & scrlTotalTasks.Value
    
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub scrlAmount_Change()
    fraAmount.Caption = "Quantidade: " & scrlAmount.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Amount = scrlAmount.Value
End Sub

Private Sub scrlNPC_Change()
    If Not scrlNPC.Value = 0 Then
        If Not Trim(NPC(scrlNPC.Value).Name) = vbNullString Then
            fraTaskNPC.Caption = "NPC: " & Trim$(NPC(scrlNPC.Value).Name)
        Else
            fraTaskNPC.Caption = "NPC: (Sem nome)ID: " & scrlNPC.Value
        End If
    Else
        fraTaskNPC.Caption = "NPC: Nenhum"
    End If
    
    Quest(EditorIndex).Task(scrlTotalTasks.Value).NPC = scrlNPC.Value
End Sub

Private Sub scrlItem_Change()
    If Not scrlItem.Value = 0 Then
        If Not Trim(Item(scrlItem.Value).Name) = vbNullString Then
            fraTaskItem.Caption = "Item: " & Trim$(Item(scrlItem.Value).Name)
        Else
            fraTaskItem.Caption = "Item: (Sem nome)ID: " & scrlItem.Value
        End If
    Else
        fraTaskItem.Caption = "Item: Nenhum"
    End If
    
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Item = scrlItem.Value
End Sub
