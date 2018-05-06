Attribute VB_Name = "modQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2
Public Const EDITOR_TASKS As Byte = 7

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

Public Quest_Changed(1 To MAX_QUESTS) As Boolean

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

Public PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec

Private Type PlayerQuestRec
    Status As Byte
    ActualTask As Byte
    CurrentCount As Byte 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Integer
    Value As Integer
End Type

Private Type QuestGiveItemRec
    Item As Integer
    Value As Integer
End Type

Private Type QuestTakeItemRec
    Item As Integer
    Value As Integer
End Type

Private Type QuestRewardItemRec
    Item As Integer
    Value As Integer
End Type
'/Alatar v1.2

Private Type TaskRec
    Order As Byte
    NPC As Integer
    Item As Integer
    Map As Byte
    Resource As Byte
    Amount As Integer
    TaskLog As String * 100
    QuestEnd As Boolean
End Type

' QuestAlatar Modified by Dooolly
Private Type QuestRec
    Name As String * 30
    Repeat As Byte
    StartMessage As String * 255
    MidMessage As String * 255
    FinishMessage As String * 255

    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
    
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredClass(1 To 5) As Long
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
    
    RewardExp As Long
    RewardMoney As Long
    RewardTitle As Byte
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
    
    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2
End Type

' ////////////
' // Editor //
' ////////////

Public Sub QuestEditorInit()
Dim i As Long
    
    If frmEditor_Quest.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1

    With frmEditor_Quest
        'Alatar v1.2
        .txtName = Trim$(Quest(EditorIndex).Name)
        
        If Quest(EditorIndex).Repeat = 1 Then
            .chkRepeat.Value = 1
        Else
            .chkRepeat.Value = 0
        End If
        
        .txtStartMessage = Trim$(Quest(EditorIndex).StartMessage)
        .txtMidMessage = Trim$(Quest(EditorIndex).MidMessage)
        .txtFinishMessage = Trim$(Quest(EditorIndex).FinishMessage)
        
        .scrlReqLevel.Value = Quest(EditorIndex).RequiredLevel
        .scrlReqQuest.Value = Quest(EditorIndex).RequiredQuest
        .scrlExp.Value = Quest(EditorIndex).RewardExp
        .scrlRewardMoney.Value = Quest(EditorIndex).RewardMoney
        .scrlRewardTitle.Value = Quest(EditorIndex).RewardTitle
        .scrlReqClass.Value = 1
        .scrlReqItem.Value = 1
        .scrlReqItemValue.Value = 1
        .scrlItemRew.Value = 1
        .scrlItemRewValue.Value = 1
        
        'Update the lists
        UpdateQuestGiveItems
        UpdateQuestTakeItems
        UpdateQuestRewardItems
        UpdateQuestRequirementItems
        UpdateQuestClass
        
        '/Alatar v1.2
        
        'load task nº1
        .scrlTotalTasks.Value = 1
        LoadTask EditorIndex, 1
        
    End With

    Quest_Changed(EditorIndex) = True
    
End Sub

'Alatar v1.2
Public Sub UpdateQuestGiveItems()
    Dim i As Byte
    
    frmEditor_Quest.lstGiveItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).GiveItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstGiveItem.AddItem "-"
            Else
                frmEditor_Quest.lstGiveItem.AddItem Trim$(Trim$(Item(.Item).Name) & ": " & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestTakeItems()
    Dim i As Long
    
    frmEditor_Quest.lstTakeItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).TakeItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstTakeItem.AddItem "-"
            Else
                frmEditor_Quest.lstTakeItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRewardItems()
    Dim i As Long
    
    frmEditor_Quest.lstItemRew.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RewardItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstItemRew.AddItem "-"
            Else
                frmEditor_Quest.lstItemRew.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestRequirementItems()
    Dim i As Long
    
    frmEditor_Quest.lstReqItem.Clear
    
    For i = 1 To MAX_QUESTS_ITEMS
        With Quest(EditorIndex).RequiredItem(i)
            If .Item = 0 Then
                frmEditor_Quest.lstReqItem.AddItem "-"
            Else
                frmEditor_Quest.lstReqItem.AddItem Trim$(Trim$(Item(.Item).Name) & ":" & .Value)
            End If
        End With
    Next
End Sub

Public Sub UpdateQuestClass()
    Dim i As Long
    
    frmEditor_Quest.lstReqClass.Clear
    
    For i = 1 To Max_Classes
        If Quest(EditorIndex).RequiredClass(i) = 0 Then
            frmEditor_Quest.lstReqClass.AddItem "-"
        Else
            frmEditor_Quest.lstReqClass.AddItem Trim$(Trim$(Class(Quest(EditorIndex).RequiredClass(i)).Name))
        End If
    Next
End Sub
'/Alatar v1.2

Public Sub QuestEditorOk()
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
End Sub

Public Sub QuestEditorCancel()
    Editor = 0
    Unload frmEditor_Quest
    ClearChanged_Quest
    ClearQuests
    SendRequestQuests
End Sub

Public Sub ClearChanged_Quest()
    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
End Sub

' //////////////
' // DATABASE //
' //////////////

Sub ClearQuest(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
End Sub

Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong CSaveQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Sub SendRequestQuests()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestQuests
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub UpdateQuestLog()
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuestLogUpdate
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub PlayerHandleQuest(ByVal QuestNum As Long, ByVal Order As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CPlayerHandleQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteLong Order '1=accept quest, 2=cancel quest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ///////////////
' // Functions //
' ///////////////

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal QuestNum As Integer) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If PlayerQuest(QuestNum).Status = QUEST_STARTED Then 'Status=1 means started
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal QuestNum As Integer) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

Public Function GetQuestNum(ByVal QuestName As String) As Integer
    Dim i As Long
    GetQuestNum = 0
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) = Trim$(QuestName) Then
            GetQuestNum = i
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

'Subroutine that load the desired task in the form
Public Sub LoadTask(ByVal QuestNum As Long, ByVal TaskNum As Long)
    Dim TaskToLoad As TaskRec
    TaskToLoad = Quest(QuestNum).Task(TaskNum)
    
    With frmEditor_Quest
        'Load the task type
'        .optTask(TaskToLoad.Order).Value = True
        'Load textboxes
        .txtTaskLog.Text = "" & Trim$(TaskToLoad.TaskLog)
        'Set scrolls to 0 and disable them so they can be enabled when needed
        .scrlNPC.Value = 0
        .scrlItem.Value = 0
        .scrlMap.Value = 0
        .scrlResource.Value = 0
        .scrlAmount.Value = 1
        .fraTaskNPC.Visible = False
        .fraTaskItem.Visible = False
        .fraTaskItem.Top = .fraTaskNPC.Top
        .scrlMap.Enabled = False
        .scrlResource.Enabled = False
        .fraAmount.Visible = False
        
        
        If TaskToLoad.QuestEnd = True Then
            .chkEnd.Value = 1
        Else
            .chkEnd.Value = 0
        End If
        
        Select Case TaskToLoad.Order
            Case 0 'Nothing
                
            Case QUEST_TYPE_GOSLAY '1
                .fraTaskNPC.Visible = True
                .scrlNPC.Value = TaskToLoad.NPC
                
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOGATHER '2
                .fraTaskItem.Visible = True
                .scrlItem.Value = TaskToLoad.Item
                
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTALK '3
                .fraTaskNPC.Visible = True
                .scrlNPC.Value = TaskToLoad.NPC
                
            Case QUEST_TYPE_GOREACH '4
                .scrlMap.Enabled = True
                .scrlMap.Value = TaskToLoad.Map
            
            Case QUEST_TYPE_GOGIVE '5
                .fraTaskItem.Visible = True
                .fraTaskItem.Top = .fraTaskItem.Top + .fraTaskNPC.Height
                .scrlItem.Value = TaskToLoad.Item
                
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
                .fraTaskNPC.Visible = True
                .scrlNPC.Value = TaskToLoad.NPC
                
            Case QUEST_TYPE_GOKILL '6
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOTRAIN '7
                .scrlResource.Enabled = True
                .scrlResource.Value = TaskToLoad.Resource
                
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
            Case QUEST_TYPE_GOGET '8
            
                .fraTaskNPC.Visible = True
                .scrlNPC.Value = TaskToLoad.NPC
                
                .fraTaskItem.Visible = True
                .fraTaskItem.Top = .fraTaskItem.Top + .fraTaskNPC.Height
                .scrlItem.Value = TaskToLoad.Item
                
                .fraAmount.Visible = True
                .scrlAmount.Value = TaskToLoad.Amount
                
        End Select
    End With
End Sub

' ////////////////////////
' // Visual Interaction //
' ////////////////////////

Public Sub LoadQuestlogBox(ByVal ButtonPressed As Integer)
    Dim QuestNum As Long, i As Long
    Dim QuestSay As String
    
    With frmMain
        Select Case ButtonPressed
            Case 1 'Actual Task
                .lblQuestSubtitle = "Actual Task [" + Trim$(PlayerQuest(QuestNum).ActualTask) + "]"
                If QuestCompleted(QuestNum) = False Then
                    .lblQuestSay = Trim$(Quest(QuestNum).Task(PlayerQuest(QuestNum).ActualTask).TaskLog)
                Else
                    .lblQuestSay = "."
                End If
                
            Case 2 'Last Speech
                .lblQuestSubtitle = "Last Speech"
                If PlayerQuest(QuestNum).ActualTask > 1 Then
                    '.lblQuestSay = Trim$(Quest(QuestNum).Task(PlayerQuest(QuestNum).ActualTask - 1).Speech)
                    If .lblQuestSay = "" Then

                    End If
                Else

                End If
            
            Case 3 'Quest Status
                .lblQuestSubtitle = "Quest Status"
                If PlayerQuest(QuestNum).Status = QUEST_STARTED Then
                    .lblQuestSay = "Quest in Progress. Step " & PlayerQuest(QuestNum).ActualTask & "."
                    .lblQuestExtra = "Cancel Quest"
                    .lblQuestExtra.Visible = True
                ElseIf QuestCompleted(QuestNum) Then
                    .lblQuestSay = "Completed"
                End If
                
            Case 4 'Quest Log (Main Task)
                .lblQuestSubtitle = "Main Task"
                .lblQuestSay = Trim$(Quest(QuestNum).StartMessage)
            
            Case 5 'Requirements
                .lblQuestSubtitle = "Requirements"
                QuestSay = "Level: "
                If Quest(QuestNum).RequiredLevel > 0 Then
                    QuestSay = QuestSay & "" & Quest(QuestNum).RequiredLevel & vbNewLine & "Quest: "
                Else
                    QuestSay = QuestSay & " None." & vbNewLine & "Quest: "
                End If
                If Quest(QuestNum).RequiredQuest > 0 Then
                    QuestSay = QuestSay & "" & Trim$(Quest(Quest(QuestNum).RequiredQuest).Name) & vbNewLine & "Class: "
                Else
                    QuestSay = QuestSay & " None." & vbNewLine & "Class: "
                End If
                For i = 1 To 5
                    If Quest(QuestNum).RequiredClass(i) > 0 Then
                        QuestSay = QuestSay & Trim$(Class(Quest(QuestNum).RequiredClass(i)).Name) & ". "
                    End If
                Next
                QuestSay = QuestSay & vbNewLine & "Items:"
                For i = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RequiredItem(i).Item > 0 Then
                        QuestSay = QuestSay & " " & Trim$(Item(Quest(QuestNum).RequiredItem(i).Item).Name) & "(" & Trim$(Quest(QuestNum).RequiredItem(i).Value) & ")"
                    End If
                Next
                .lblQuestSay = QuestSay
            
            Case 6 'Rewards
                .lblQuestSubtitle = "Rewards"
                QuestSay = "Experience: " & Quest(QuestNum).RewardExp & vbNewLine & "Items:"
                For i = 1 To MAX_QUESTS_ITEMS
                    If Quest(QuestNum).RewardItem(i).Item > 0 Then
                        QuestSay = QuestSay & " " & Trim$(Item(Quest(QuestNum).RewardItem(i).Item).Name) & "(" & Trim$(Quest(QuestNum).RewardItem(i).Value) & ")"
                    End If
                Next
                .lblQuestSay = QuestSay
            
            Case Else
                Exit Sub
        End Select
        
        .lblQuestName = Trim$(Quest(QuestNum).Name)
        .picQuestDialogue.Visible = True
        
    End With
End Sub

Public Sub RunQuestDialogueExtraLabel()
    If frmMain.lblQuestExtra = "Cancel Quest" Then
        PlayerHandleQuest GetQuestNum(Trim$(frmMain.lblQuestName.Caption)), 2
        frmMain.lblQuestExtra = "Extra"
        frmMain.lblQuestExtra.Visible = False
        frmMain.picQuestDialogue.Visible = False
    End If
End Sub
