VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paniel do Administrador"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEditors 
      Caption         =   "Editores:"
      ForeColor       =   &H80000011&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmbAnims 
         Caption         =   "Animações"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEditor 
         Caption         =   "Missões"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Painel de Administrador

'Copyright (C) 2007 Free Software Foundation, Inc. <http://fsf.org/>
' Everyone is permitted to copy and distribute verbatim copies
' of this license document, but changing it is not allowed.

Option Explicit

Private Sub cmbAnims_Click()
    Dim i As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditor_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Verificar se o jogador é administrador
    If GetPlayerAccess(MyIndex) <= ADMIN_DEVELOPER Then
        Me.Visible = False
        Exit Sub
    End If
    
    ' Solicitar editor
    Call RequestEditor(Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub RequestEditor(EditorIndex As Integer)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditor
    Buffer.WriteInteger EditorIndex
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
