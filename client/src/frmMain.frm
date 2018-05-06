VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picQuestDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2760
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   97
      Top             =   2400
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Label lblQuestSay 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1125
         Left            =   240
         TabIndex        =   103
         Top             =   720
         Width           =   4425
      End
      Begin VB.Label lblQuestAccept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accept Quest"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   210
         Left            =   240
         TabIndex        =   102
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblQuestClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Left            =   4200
         TabIndex        =   101
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblQuestName 
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   100
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblQuestExtra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   99
         Top             =   1920
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblQuestSubtitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Subtitle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   98
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.PictureBox picEventChat 
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   -240
      ScaleHeight     =   1800
      ScaleWidth      =   7140
      TabIndex        =   90
      Top             =   3720
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 4 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   4
         Left            =   5280
         TabIndex        =   96
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 3 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   3
         Left            =   3600
         TabIndex        =   95
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 2 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   1920
         TabIndex        =   94
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 1 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   93
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblEventChatContinue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Continue >"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   92
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblEventChat 
         BackColor       =   &H000C0E10&
         Caption         =   "This is text that appears for an event."
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   0
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   69
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   68
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   45
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   73
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   72
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   71
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5B5B5&
      ForeColor       =   &H80000008&
      Height          =   8010
      Left            =   12000
      ScaleHeight     =   532
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2865
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "Screenshot Map"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   7560
         Width           =   2295
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   6720
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   27
         Top             =   6360
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   25
         Top             =   5760
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   16
         X2              =   168
         Y1              =   472
         Y2              =   472
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   168
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   304
         Y2              =   304
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Editors:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   144
         Y2              =   144
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   2865
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   77
      Top             =   11880
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3285
         TabIndex        =   82
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3375
         TabIndex        =   79
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   78
         Top             =   1560
         Width           =   285
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   38
      Top             =   9720
      Visible         =   0   'False
      Width           =   7140
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   40
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3240
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3300
         TabIndex        =   41
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6480
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   56
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7680
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   74
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   12360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   76
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   12000
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   70
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   1680
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   44
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2715
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1545
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   375
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   57
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3855
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   59
         Top             =   465
         Width           =   2895
      End
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   58
         Top             =   465
         Width           =   2895
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   61
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   60
         Top             =   4500
         Width           =   1815
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   55
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   83
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   89
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   87
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   86
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   85
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   84
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   46
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   52
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   53
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   49
         Top             =   840
         Width           =   1935
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   51
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   47
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.Label lblEXP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   67
      Top             =   1080
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   66
      Top             =   750
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   65
      Top             =   420
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image imgEXPBar 
      Height          =   240
      Left            =   7770
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Image imgMPBar 
      Height          =   240
      Left            =   7770
      Top             =   750
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Image imgHPBar 
      Height          =   240
      Left            =   7770
      Top             =   420
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblPing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   64
      Top             =   1920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblGold 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0g"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   63
      Top             =   1515
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   6
      Left            =   10245
      Top             =   3450
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   5
      Left            =   9045
      Top             =   3450
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   7845
      Top             =   3450
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   10245
      Top             =   2850
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   9045
      Top             =   2850
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   7845
      Top             =   2850
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub Form_DblClick()
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DblClick_Handle
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' move GUI
    picAdmin.Left = 544
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
    Else
        MouseDown_Handle Button, Shift, X, Y
    End If
    
    If frmEditor_Events.Visible Then frmEditor_Events.SetFocus

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    
    MouseMove_Handle Button, Shift, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblChoices_Click(Index As Integer)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong Index
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    txtCurrency.Text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgLeaveShop_Click()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CCloseShop
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    picCover.Visible = False
    picShop.Visible = False
    InShop = 0
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsNumeric(txtCurrency.Text) Then
        If Val(txtCurrency.Text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.Text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
        Select Case CurrencyMenu
            Case 1 ' drop item
                SendDropItem tmpCurrencyItem, Val(txtCurrency.Text)
            Case 2 ' deposit item
                DepositItem tmpCurrencyItem, Val(txtCurrency.Text)
            Case 3 ' withdraw item
                WithdrawItem tmpCurrencyItem, Val(txtCurrency.Text)
            Case 4 ' offer trade item
                TradeItem tmpCurrencyItem, Val(txtCurrency.Text)
        End Select
    Else
        AddText "Please enter a valid amount.", BrightRed
        Exit Sub
    End If
    
    picCurrency.Visible = False
    tmpCurrencyItem = 0
    txtCurrency.Text = vbNullString
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "Click on the item in the shop you wish to buy.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    AddText "Double-click on the item in your inventory you wish to sell.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDialogue_Button_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the handler
    dialogueHandler Index
    
    picDialogue.Visible = False
    dialogueIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub lblEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub lblEventChatContinue_Click()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong 0
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblEventChatContinue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub lblEventChatContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblEventChatContinue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
        SendPartyRequest
    Else
        AddText "Invalid invitation target.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Party.Leader > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblTrainStat_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'ALATAR

'QuestDialogue:

Private Sub lblQuestAccept_Click()
    PlayerHandleQuest CLng(lblQuestAccept.Tag), 1
    picQuestDialogue.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
End Sub

Private Sub lblQuestExtra_Click()
    RunQuestDialogueExtraLabel
End Sub

Private Sub lblQuestClose_Click()
    picQuestDialogue.Visible = False
    lblQuestExtra.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
End Sub

Private Sub optMOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 0
    ' stop music playing
    StopMusic
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOn_Click()
Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 1
    ' start music playing
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopAllSounds
    Options.sound = 0
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.sound = 1
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCover_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub picEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MouseUp_Handle Button, Shift, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim TempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With TempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopItem = IsShopItem(X, Y)
    
    If shopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(shopItem)
                    AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", White
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem shopItem
        End Select
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopslot As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    shopslot = IsShopItem(X, Y)

    If shopslot <> 0 Then
        x2 = X + picShop.Left + picShopItems.Left + 1
        y2 = Y + picShop.Top + picShopItems.Top + 1
        UpdateDescWindow Shop(InShop).TradeItem(shopslot).Item, x2, y2
        LastItemDesc = Shop(InShop).TradeItem(shopslot).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
Dim SpellNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellNum = IsPlayerSpell(SpellX, SpellY)

    If SpellNum <> 0 Then
        Call CastSpell(SpellNum)
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).Num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).Num) ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateDescWindow TradeTheirOffer(TradeNum).Num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).Num ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.Value).Name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                picAdmin.Visible = Not picAdmin.Visible
            End If
    End Select
    
    ' hotbar
    For i = 1 To MAX_HOTBAR
        If KeyCode = 48 + i Then
           SendHotbarUse i
        ElseIf KeyCode = 48 Then
           SendHotbarUse 10
           Exit For
        End If
    Next
    
    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With TempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i).Num > 0 And PlayerSpells(i).Num <= MAX_SPELLS Then

            With TempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim TempRec As RECT
    Dim i As Long
    Dim itemNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)
        Else
            itemNum = TradeTheirOffer(i).Num
        End If

        If itemNum > 0 And itemNum <= MAX_ITEMS Then

            With TempRec
                .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.Text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtASprite.Text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.Text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.Text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.Text)) Or Not IsNumeric(Trim$(txtAAccess.Text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.Text), CLng(Trim$(txtAAccess.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' bank
Private Sub picBank_DblClick()
Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
         If GetBankItemNum(bankNum) = ITEM_TYPE_NONE Then Exit Sub
         
             If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 3 ' withdraw
                lblCurrency.Caption = "How many do you want to withdraw?"
                tmpCurrencyItem = bankNum
                txtCurrency.Text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
         WithdrawItem bankNum, 0
         Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long
                        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(X, Y)
    
    If bankNum <> 0 Then
        
        If Button = 1 Then
            DragBankSlotNum = bankNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim Rec_Pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With Rec_Pos
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= Rec_Pos.Left And X <= Rec_Pos.Right Then
                If Y >= Rec_Pos.Top And Y <= Rec_Pos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long, itemNum As Long, ItemType As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y
    
    If DragBankSlotNum > 0 Then
        With frmMain.picTempBank
            .Top = Y + picBank.Top
            .Left = X + picBank.Left
            .Visible = True
            .ZOrder (0)
        End With
    Else
        bankNum = IsBankItem(X, Y)
        
        If bankNum <> 0 Then
            
            x2 = X + picBank.Left + 1
            y2 = Y + picBank.Top + 1
            UpdateDescWindow Bank.Item(bankNum).Num, x2, y2
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
Dim TempRec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
        
            With TempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
